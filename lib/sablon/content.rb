module Sablon
  module Content
    class << self
      def wrap(value)
        case value
        when Sablon::Content
          value
        else
          if type = type_wrapping(value)
            type.new(value)
          else
            raise ArgumentError, "Could not find Sablon content type to wrap #{value.inspect}"
          end
        end
      end

      def make(type_id, *args)
        if types.key?(type_id)
          types[type_id].new(*args)
        else
          raise ArgumentError, "Could not find Sablon content type with id '#{type_id}'"
        end
      end

      def register(content_type)
        types[content_type.id] = content_type
      end

      def remove(content_type_or_id)
        types.delete_if {|k,v| k == content_type_or_id || v == content_type_or_id }
      end

      private
      def type_wrapping(value)
        types.values.reverse.detect { |type| type.wraps?(value) }
      end

      def types
        @types ||= {}
      end
    end

    class Image < Struct.new(:source, :data)
      require 'image_spec'

      def initialize(source)
        self.source = source
        case source
        when ::File, ::Tempfile, ::StringIO
          self.data = source
        when ::Pathname, ::String
          self.data = File.open(source, 'rb')
        end
      end

      def spec
        ret = ImageSpec.new(data)
        data.rewind if data.is_a?(::StringIO)
        ret
      end

      def self.id
        :image
      end

      def self.wraps?(source)
        case source
        when ::String
          self.wraps?(Pathname.new(source))
        when ::Pathname
          source.exist? && source.file? && self.wraps?(File.open(source))
        when ::File, ::Tempfile, ::StringIO
          ImageSpec.new(source)
        end
      rescue ImageSpec::Error
        false
      end

      def append_to(paragraph, display_node)
        false
      end
    end

    class String < Struct.new(:string)
      include Sablon::Content
      def self.id; :string end
      def self.wraps?(value)
        value.respond_to?(:to_s)
      end

      def initialize(value)
        super value.to_s
      end

      def append_to(paragraph, display_node)
        string.scan(/[^\n\r]+|[\n\r]+/).reverse.each do |part|
          if part[/[\n\r]/]
            display_node.add_next_sibling Nokogiri::XML::Node.new "w:br", display_node.document
          else
            text_part = display_node.dup
            text_part.content = part
            display_node.add_next_sibling text_part
          end
        end
      end
    end

    class WordML < Struct.new(:xml)
      include Sablon::Content
      def self.id; :word_ml end
      def self.wraps?(value) false end

      def append_to(paragraph, display_node)
        Nokogiri::XML.fragment(xml).children.reverse.each do |child|
          paragraph.add_next_sibling child
        end
        paragraph.remove
      end
    end

    class InlineWordML < Struct.new(:xml, :replace_run)
      include Sablon::Content
      def self.id; :inline_word_ml end
      def self.wraps?(value) false end

      def append_to(paragraph, display_node)
        if replace_run
          display_node = display_node.at_xpath('./ancestor::w:r')
        end

        Nokogiri::XML.fragment(xml).children.reverse.each do |child|
          display_node.add_next_sibling child
        end

        display_node.remove if replace_run
      end
    end

    class Markdown < Struct.new(:word_ml)
      include Sablon::Content
      def self.id; :markdown end
      def self.wraps?(value) false end

      def initialize(markdown)
        warn "[DEPRECATION] `Sablon::Content::Markdown` is deprecated.  Please use `Sablon::Content::HTML` instead."
        redcarpet = ::Redcarpet::Markdown.new(::Redcarpet::Render::HTML.new)
        word_ml = Sablon.content(:html, redcarpet.render(markdown))
        super word_ml
      end

      def append_to(*args)
        word_ml.append_to(*args)
      end
    end

    class HTML < Struct.new(:word_ml, :html, :styles)
      include Sablon::Content
      def self.id; :html end
      def self.wraps?(value) false end

      def initialize(html, styles = nil)
        self.html = html
        self.styles = styles
        @converter = HTMLConverter.new
        @converter.styles = styles
      end

      def numbering=(numbering)
        @converter.numbering = numbering
      end

      def word_ml
        Sablon.content(:word_ml, @converter.process(self.html))
      end

      def present?
        Nokogiri::HTML(html.to_s).content.present?
      end

      def append_to(*args)
        word_ml.append_to(*args)
      end
    end

    register Sablon::Content::String
    register Sablon::Content::WordML
    register Sablon::Content::InlineWordML
    register Sablon::Content::Markdown
    register Sablon::Content::HTML
    register Sablon::Content::Image
  end
end
