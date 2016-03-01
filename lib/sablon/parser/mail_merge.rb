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
        PICTURE_NS_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
        MAIN_NS_URI = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        KEY_PATTERN = /^(=[^ ]+)/

        attr_accessor :resources

        def initialize(node, resources)
          self.resources = resources
          @node = node
          @raw_expression = node.at_xpath('.//pic:cNvPr', 'pic' => PICTURE_NS_URI)['name'].strip
          binding.pry unless valid?
        end


        def merge_field_nodes
          @nodes.find{ |el| el.is_a?(Array) && el.take_while(&method(:not_separator?)).join(' ').strip[KEY_PATTERN] }
        end

        def replace(content)
          if content.is_a?(Content::Image) && node = @node
            image_id = next_image_id
            replace_image(node, image_id)
            resources[image_id] = content
          end
        end

        def replace_image(node, image_id)
          node.at_xpath('.//a:blip', 'a' => MAIN_NS_URI)['r:embed'] = image_id
        end

        def next_image_id
          rid = resources.keys.map{ |rid| (rid[/^rId\d+$/] && rid.sub('rId', '')).to_i }.max.to_i + 1
          "rId#{ rid }"
        end

        def remove
          false
        end

        def expression
          $1 if @raw_expression =~ KEY_PATTERN
        end

        class << self
          def valid_candidate?(node)
            return false if node.name != 'drawing'
            prop = node.at_xpath('.//pic:cNvPr', 'pic' => PICTURE_NS_URI)
            prop['name'].strip[/^=/]
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
          elsif node.name == 'drawing'
            field = IncludeImage.new(node, resources) if IncludeImage.valid_candidate?(node)
          elsif node.name == "fldChar" && node["w:fldCharType"] == "begin"
            field = build_complex_field(node)
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
