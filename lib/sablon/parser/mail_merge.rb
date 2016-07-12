require 'fileutils'

module Sablon
  module Parser
    class MailMerge
      attr_accessor :resources, :numbering

      def initialize(resources, numbering)
        self.resources = resources
        self.numbering = numbering
      end

      class MergeField
        attr_accessor :nodes
        KEY_PATTERN = /^\s*MERGEFIELD\s*([\s\S]+?)\s*\\\*\s*MERGEFORMAT\s*$/

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

        def self.from_cache(all_nodes, cache)
          field = allocate
          field.resources = cache[:resources]
          field.instance_variable_set(:@raw_expression, cache[:raw_expression])
          field.instance_variable_set(:@node, all_nodes[cache[:node]])
          field
        end

        def to_cache(all_nodes)
          {
            type: self.class.name,
            node: all_nodes.index(@node),
            raw_expression: @raw_expression
          }
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
            return false if node.at_xpath('.//w:drawing')
            prop = node.at_xpath('.//pic:cNvPr', 'pic' => PICTURE_NS_URI)
            prop && prop['name'].strip[/^=/]
          end
        end
      end

      class ComplexField < MergeField
        def initialize(nodes)
          @nodes = nodes
          @raw_expression = @nodes.flatten(1).take_while{ |n| n != separate_node }.map {|n| n.search(".//w:instrText").map(&:content) }.join
        end

        def self.from_cache(all_nodes, cache)
          ancestor_paragraph_set = Nokogiri::XML::NodeSet.new(all_nodes.first.document, all_nodes.values_at(*cache[:ancestor_paragraphs]))
          ancestor_row_set = Nokogiri::XML::NodeSet.new(all_nodes.first.document, all_nodes.values_at(*cache[:ancestor_rows]))
          ancestor_set = Nokogiri::XML::NodeSet.new(all_nodes.first.document, all_nodes.values_at(*cache[:ancestors]))

          field = allocate
          field.instance_variable_set(:@raw_expression, cache[:raw_expression])
          field.instance_variable_set(:@nodes, all_nodes.values_at(*cache[:nodes]))
          field.instance_variable_set(:@ancestors, ancestor_set)
          field.instance_variable_set(:@ancestor_paragraphs, ancestor_paragraph_set)
          field.instance_variable_set(:@ancestor_rows, ancestor_row_set)

          binding.pry if Object.instance_variable_defined?(:@pryme)
          field
        end

        def to_cache(all_nodes)
          ancestor_paragraphs = ancestors(".//w:p")
          ancestor_rows = ancestors(".//w:tr")
          all_ancestors = ancestors
          {
            type: self.class.name,
            ancestors: all_ancestors.is_a?(Nokogiri::XML::NodeSet) ? all_ancestors.map(&all_nodes.method(:index)) : all_ancestors,
            ancestor_paragraphs: ancestor_paragraphs.is_a?(Nokogiri::XML::NodeSet) ? ancestor_paragraphs.map(&all_nodes.method(:index)) : ancestor_paragraphs,
            ancestor_rows: ancestor_rows.is_a?(Nokogiri::XML::NodeSet) ? ancestor_rows.map(&all_nodes.method(:index)) : ancestor_rows,
            nodes: @nodes.map(&all_nodes.method(:index)),
            raw_expression: @raw_expression
          }
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
          if (args.first == ".//w:p") && @ancestor_paragraphs
            @ancestor_paragraphs
          elsif (args.first == ".//w:tr") && @ancestor_rows
            @ancestor_rows
          elsif @ancestors && !args.any?
            @ancestors
          else
            @nodes.first.ancestors(*args)
          end
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
          return @separate_node if defined? @separate_node
          @separate_node = @nodes.detect {|n| !n.search(".//w:fldChar[@w:fldCharType='separate']").empty? }
        end
      end

      class SimpleField < MergeField
        def initialize(node)
          @node = node
          @raw_expression = @node["w:instr"]
        end

        def to_cache(all_nodes)
          ancestor_paragraphs = ancestors(".//w:p")
          ancestor_rows = ancestors(".//w:tr")
          all_ancestors = ancestors
          {
            type: self.class.name,
            ancestors: all_ancestors.is_a?(Nokogiri::XML::NodeSet) ? all_ancestors.map(&all_nodes.method(:index)) : all_ancestors,
            ancestor_paragraphs: ancestor_paragraphs.is_a?(Nokogiri::XML::NodeSet) ? ancestor_paragraphs.map(&all_nodes.method(:index)) : ancestor_paragraphs,
            ancestor_rows: ancestor_rows.is_a?(Nokogiri::XML::NodeSet) ? ancestor_rows.map(&all_nodes.method(:index)) : ancestor_rows,
            node: all_nodes.index(@node),
            raw_expression: @raw_expression
          }
        end

        def self.from_cache(all_nodes, cache)
          ancestor_paragraph_set = Nokogiri::XML::NodeSet.new(all_nodes.first.document, all_nodes.values_at(*cache[:ancestor_paragraphs]))
          ancestor_row_set = Nokogiri::XML::NodeSet.new(all_nodes.first.document, all_nodes.values_at(*cache[:ancestor_rows]))
          ancestor_set = Nokogiri::XML::NodeSet.new(all_nodes.first.document, all_nodes.values_at(*cache[:ancestors]))

          field = allocate
          field.instance_variable_set(:@raw_expression, cache[:raw_expression])
          field.instance_variable_set(:@node, all_nodes[cache[:node]])
          field.instance_variable_set(:@ancestors, ancestor_set)
          field.instance_variable_set(:@ancestor_paragraphs, ancestor_paragraph_set)
          field.instance_variable_set(:@ancestor_rows, ancestor_row_set)
          field
        end

        def replace(content)
          replace_field_display(@node, content)
          @node.replace(@node.children)
        end

        def remove
          @node.remove
        end

        def ancestors(*args)
          if (args.first == ".//w:p") && @ancestor_paragraphs
            @ancestor_paragraphs
          elsif (args.first == ".//w:tr") && @ancestor_rows
            @ancestor_rows
          elsif @ancestors && !args.any?
            @ancestors
          else
            @node.ancestors(*args)
          end
        end

        def start_node
          @node
        end

        alias_method :end_node, :start_node
      end

      def cache_file(xml_id)
        raise 'no cache dir provided' unless Sablon.cache_dir
        ret = Pathname.new(Sablon.cache_dir).join("#{ xml_id }.dump")
        FileUtils.mkdir_p ret.dirname
        ret
      end

      def write_cache(xml_id, field_cache)
        File.open(cache_file(xml_id), File::RDWR|File::CREAT, 0644) {|f|
          f.flock(File::LOCK_EX)
          f.binmode
          f.write(Marshal.dump(field_cache))
        }
      end

      def read_cache(xml_id)
        File.open(cache_file(xml_id), "r") do |f|
          f.flock(File::LOCK_SH)
          f.binmode
          Marshal.load(f.read)
        end
      rescue Errno::ENOENT, ArgumentError
        nil
      end

      def parse_fields(xml)
        xml_id = "#{ Sablon::VERSION }::#{ Digest::SHA512.hexdigest(xml.to_s) }"
        all_nodes = xml.enum_for(:traverse).to_a
        # all_nodes = xml.xpath("/descendant-or-self::node()").to_a
        cached_fields = read_cache(xml_id)

        if cached_fields
          fields = cached_fields.map do |cache|
            self.class.const_get(cache[:type]).from_cache(
              all_nodes,
              cache.merge(resources: resources)
            )
          end
        else
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

          cached_fields = fields.map{ |f| f.to_cache(all_nodes) }
          write_cache(xml_id, cached_fields)
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
