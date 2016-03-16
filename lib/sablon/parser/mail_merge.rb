module Sablon
  module Parser
    class MailMerge
      attr_accessor :resources

      def initialize(resources)
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
          paragraph = node.ancestors(".//w:p").first || node.ancestors.search('tmp').try(:first)
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

        class SizeCalculator < Struct.new(:node, :resources, :x, :y, :image_x, :image_y, :offx, :offy, :realx, :realy)
          require 'yaml'
          def initialize(node, realx, realy, resources)
            self.realx = realx
            self.realy = realy
            self.node = node
            self.x = extent_node['cx'].to_i
            self.y = extent_node['cy'].to_i
            self.image_x = shape_extent_node['cx'].to_i
            self.image_y = shape_extent_node['cy'].to_i
            self.offx = src_rect_node['r'].to_i
            self.offy = src_rect_node['b'].to_i
            self.resources = resources


            configuration['width'] ||= 'rect'
            configuration['height'] ||= 'rect'
          end

          def update_id
            newid = node.document.root.search('.//w:drawing//wp:docPr').map{ |pr| pr['id'].to_i }.max + resources.size
            property_node['id'] = newid
          end

          def apply
            extent_node['cx'] = newx
            extent_node['cy'] = newy

            shape_extent_node['cx'] = newx
            shape_extent_node['cy'] = newy

            stretch = stretch_node
            stretch.children.map(&:remove)

            update_id

            src_rect_node['l'] = format_percent(new_off_x[0].to_f / new_image_x.to_f) if new_off_x[0] != 0
            src_rect_node['r'] = format_percent(new_off_x[1].to_f / new_image_x.to_f) if new_off_x[1] != 0
            src_rect_node['t'] = format_percent(new_off_y[0].to_f / new_image_y.to_f) if new_off_y[0] != 0
            src_rect_node['b'] = format_percent(new_off_y[1].to_f / new_image_y.to_f) if new_off_y[1] != 0

            property_node['descr'] =  property_node['descr'].to_s.sub(META_PATTERN, '')
          end

          def format_percent(number)
            parts = (number.to_f * 100).divmod(1)
            "#{ parts[0] }#{ (1000 * parts[1]).round.to_s.rjust(3, '0') }"
          end

          def configuration
            @config ||= (property_node['descr'].to_s =~ META_PATTERN) ? parse_config($1) : {}
          end

          def full_auto?
            ![configuration['height'], configuration['width']].include?('rect')
          end

          def newy
            configuration['height'] == 'auto' ? auto_y : y
          end

          def newx
            configuration['width'] == 'auto' ? auto_x : x
          end

          def auto_x
            if full_auto?
              if (realy.to_f / y.to_f ) <= (realx.to_f / x.to_f)
                x
              else
                ((auto_y * realx) / realy).to_i
              end
            else
              ((y * realx) / realy).to_i
            end
          end

          def auto_y
            if full_auto?
              if (realy.to_f / y.to_f ) > (realx.to_f / x.to_f)
                y
              else
                ((auto_x * realy) / realx).to_i
              end
            else
              ((x * realy) / realx).to_i
            end
          end

          def new_image_x
            configuration['width'] == 'rect' ? x : auto_x
          end

          def new_image_y
            configuration['height'] == 'rect' ? y : auto_y
          end

          def new_off_x
            {
              'rect' => [0, 0],
              'auto' => [0, 0],
              'left' => [0, 0],
              'right' => [new_image_x - newx, 0],
              'middle' => [((new_image_x - newx) / 2).to_i]*2
            }[configuration['width']]
          end

          def new_off_y
            {
              'rect' => [0,0],
              'auto' => [0,0],
              'top' => [0,0],
              'bottom' => [new_image_y - newy, 0],
              'middle' => [((new_image_y - newy) / 2).to_i]*2,
            }[configuration['height']]
          end

          private

          def src_rect_node
            ret = node.at_xpath('.//a:srcRect', 'a' => MAIN_NS_URI)
            return ret if ret

            stretch_node.before '<a:srcRect/>'
            src_rect_node
          end

          def property_node
            node.search('.//wp:docPr').first
          end

          def extent_node
            node.search('.//wp:extent').first
          end

          def shape_extent_node
            node.
              at_xpath('.//pic:spPr', 'pic' => PICTURE_NS_URI).
              at_xpath('.//a:ext', 'a' => MAIN_NS_URI)
          end

          def stretch_node
            node.
              at_xpath('.//pic:blipFill', 'pic' => PICTURE_NS_URI).
              at_xpath('.//a:stretch', 'a' => MAIN_NS_URI)
          end

          def parse_config(hash)
            indent = -1
            YAML.load(
              hash.gsub(/([\{\},]\s*)|(=>)/) do |match|
                case match[0]
                when '{'
                  indent += 1
                  "\n#{ '  ' * indent }"
                when '}'
                  indent = [0, indent - 1].max
                  "\n#{ '  ' * indent }"
                when ','
                  "\n#{ '  ' * indent }"
                when '='
                  ': '
                end
              end
            )
          rescue Psych::SyntaxError
            {}
          end
        end

        PICTURE_NS_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
        MAIN_NS_URI = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        KEY_PATTERN = /^(=[^ ]+)/
        META_PATTERN = /^\!(\{[^\}]+\})/

        attr_accessor :resources

        def initialize(node, resources)
          self.resources = resources
          @node = node
          @raw_expression = node.at_xpath('.//pic:cNvPr', 'pic' => PICTURE_NS_URI)['name'].strip
        end

        def replace(content)
          if content.is_a?(Content::Image) && node = @node
            image_id = next_image_id
            replace_image(node, image_id)
            calc = SizeCalculator.new(node, content.spec.width, content.spec.height, resources)
            calc.apply
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
          @raw_expression = @nodes.flatten(1).take_while{ |n| n != separate_node }.map {|n| n.search(".//w:instrText").map(&:content) }.join
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
