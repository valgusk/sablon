module Sablon
  module Parser
    class MailMerge
      attr_accessor :resources

      def initialize(resources = {})
        self.resources = resources
      end

      class MergeField
        KEY_PATTERN = /^\s*MERGEFIELD\s+([^ ]+)\s+\\\*\s+MERGEFORMAT\s*$/

        def valid?
          expression
        end

        def expression
          $1 if @raw_expression =~ KEY_PATTERN
        end

        private
        def replace_field_display(node, content)
          paragraph = node.ancestors(".//w:p").first
          display_node = get_display_node(node)
          content.append_to(paragraph, display_node)
          display_node.remove
        end

        def get_display_node(node)
          node.search(".//w:t").first || node.search(".//w:instrText").first
        end
      end

      # class

      class IncludeImage < MergeField
        attr_accessor :resources

        def initialize(node, resources)
          self.resources = resources
          current_node = node.parent

          collect_block = proc do
            open_node = true
            block = [current_node]

            while open_node
              if current_node = current_node.next_element
                if IncludeImage.special_char_node?(current_node, 'end')
                  block << current_node
                  open_node = false
                elsif IncludeImage.special_char_node?(current_node, 'begin')
                  block << collect_block.()
                elsif IncludeImage.special_char_node?(current_node, 'separate')
                  block << current_node
                else
                  block << current_node
                end
              else
                open_node = false
              end
            end
            block
          end

          @nodes = collect_block.()
          @raw_expression = (merge_field_expression).to_s
          binding.pry unless valid?
        end

        def merge_field_expression
          ret = merge_field_nodes.to_a.first && merge_field_nodes.take_while(&method(:not_separator?)).join(' ').strip

          unless ret
            simple = @nodes.find{ |node| node.search('w:fldSimple').any? }
            simple = simple && simple.search('fldSimple').first
            simple ||= @nodes.find{ |node| node.name == 'fldSimple' }

            ret = simple && simple['w:instr'].strip
          end

          ret
        end

        def expression
          super
        end

        def display_node
          candidate_node = @nodes[@nodes.index(&method(:separator?)) + 1]
          case candidate_node
          when Nokogiri::XML::Node
            candidate_node
          when Array
            candidate_nodes = candidate_node.flatten
            candidate_nodes.find { |node| node.search(".//w:pict").any? }
          end
        end

        def merge_field_nodes
          @nodes.find{ |el| el.is_a?(Array) && el.take_while(&method(:not_separator?)).join(' ').strip[KEY_PATTERN] }
        end

        def merge_field_display_node
          separator_i = merge_field_nodes.to_a.index(&method(:separator?))
          merge_field_nodes.to_a[separator_i + 1] if separator_i
        end

        def separator?(node)
          return false unless node.is_a? Nokogiri::XML::Node
          IncludeImage.special_char_node?(node, 'separate')
        end

        def not_separator?(node)
          !separator?(node)
        end

        def valid?
          # display_node && super
          super
        end

        def replace(content)
          if content.is_a?(Content::Image) && node = display_node
            image_id = next_image_id
            replace_image(node, image_id)
            resources[image_id] = content
          end

          remove
        end

        def replace_image(node, image_id)
          node.search('.//v:imagedata').first['r:id'] = image_id
        end

        def next_image_id
          rid = resources.keys.map{ |rid| (rid[/^rId\d+$/] && rid.sub('rId', '')).to_i }.max.to_i + 1
          "rId#{ rid }"
        end

        def remove
          (@nodes.flatten - [display_node]).each(&:remove)
        end

        class << self
          def valid_candidate?(node)
            current_node = node.parent

            if special_char_node?(current_node, 'begin')
              current_node = current_node.next_element

              while current_node && current_node.content.strip == ''
                break if special_char_node?(current_node, 'begin')
                break if special_char_node?(current_node, 'end')

                current_node = current_node.next_element
              end

            end

            (current_node && current_node.content).to_s.strip.upcase[/^INCLUDEPICTURE/]
          end

          def special_char_node?(node, type)
            node.search(".//w:fldChar[@w:fldCharType='#{ type }']").any?
          end
        end
      end

      class ComplexField < MergeField
        def initialize(nodes)
          @nodes = nodes
          @raw_expression = @nodes.flat_map {|n| n.search(".//w:instrText").map(&:content) }.join
        end

        def valid?
          separate_node && get_display_node(pattern_node) && expression
        end

        def replace(content)
          replace_field_display(pattern_node, content)
          (@nodes - [pattern_node]).each(&:remove)
        end

        def remove
          @nodes.each(&:remove)
        end

        def ancestors(*args)
          @nodes.first.ancestors(*args)
        end

        def start_node
          @nodes.first
        end

        def end_node
          @nodes.last
        end

        private
        def pattern_node
          separate_node.next_element
        end

        def separate_node
          @nodes.detect {|n| !n.search(".//w:fldChar[@w:fldCharType='separate']").empty? }
        end
      end

      class SimpleField < MergeField
        def initialize(node)
          @node = node
          @raw_expression = @node["w:instr"]
        end

        def replace(content)
          replace_field_display(@node, content)
          @node.replace(@node.children)
        end

        def remove
          @node.remove
        end

        def ancestors(*args)
          @node.ancestors(*args)
        end

        def start_node
          @node
        end
        alias_method :end_node, :start_node
      end

      def parse_fields(xml)
        fields = []
        xml.traverse do |node|
          if node.name == "fldSimple"
            field = SimpleField.new(node)
          elsif node.name == "fldChar" && node["w:fldCharType"] == "begin"
            if IncludeImage.valid_candidate?(node)
              field = IncludeImage.new(node, resources)
            else
              field = build_complex_field(node)
            end
          end
          fields << field if field && field.valid?
        end
        fields
      end

      private
      def build_complex_field(node)
        possible_field_node = node.parent
        field_nodes = [possible_field_node]
        while possible_field_node && possible_field_node.search(".//w:fldChar[@w:fldCharType='end']").empty?
          possible_field_node = possible_field_node.next_element
          field_nodes << possible_field_node
        end
        ComplexField.new(field_nodes)
      end
    end
  end
end
