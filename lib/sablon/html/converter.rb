require "sablon/html/ast"
require "sablon/html/visitor"

module Sablon
  class HTMLConverter
    attr_accessor :numbering
    attr_accessor :styles

    class ASTBuilder
      Layer = Struct.new(:items, :ilvl)

      def initialize(nodes)
        @layers = [Layer.new(nodes, false)]
        @root = Root.new([])
      end

      def to_ast
        @root
      end

      def new_layer(ilvl: false)
        @layers.push Layer.new([], ilvl)
      end

      def next
        current_layer.items.shift
      end

      def push(node)
        @layers.last.items.push node
      end

      def push_all(nodes)
        nodes.each(&method(:push))
      end

      def done?
        !current_layer.items.any?
      end

      def nested?
        ilvl > 0
      end

      def ilvl
        @layers.select { |layer| layer.ilvl }.size - 1
      end

      def emit(node)
        @root.nodes << node
      end

      private
      def current_layer
        if @layers.any?
          last_layer = @layers.last
          if last_layer.items.any?
            last_layer
          else
            @layers.pop
            current_layer
          end
        else
          Layer.new([], false)
        end
      end
    end

    def process(input)
      fail 'Numbering instance is not set!' unless numbering
      processed_ast(input).to_docx
    end

    def processed_ast(input)
      ast = build_ast(input)
      ast.accept LastNewlineRemoverVisitor.new
      ast
    end

    def build_ast(input)
      doc = Nokogiri::HTML.fragment(input)
      @builder = ASTBuilder.new(doc.children)

      while !@builder.done?
        ast_next_paragraph
      end
      @builder.to_ast
    end

    private
    def styles
      [1..6].reduce(
        div: 'Normal',
        p: 'Paragraph',
        ul: 'ListBullet',
        ol: 'ListNumber',
      ) do |memo, headingNum|
        memo.merge("h#{ headingNum }".to_sym => "Heading#{ headingNum }")
      end.merge(@styles || {})
    end

    def get_style(tag)
      styles[tag.downcase.to_sym] || 'Normal'
    end

    def ast_next_paragraph
      orphans = []
      node = @builder.next

      while node && (node.text? || ['br', 'strong', 'b', 'em', 'i', 'u'].include?(node.name))
        if !node.text?  || node.text[/\S/]
          orphans << node
        end
        node = @builder.next
      end

      if orphans.any?
        @builder.new_layer
        @builder.emit Paragraph.new(get_style('div'), ast_text(orphans))
      end

      return unless node

      if node.name == 'div'
        @builder.new_layer
        @builder.emit Paragraph.new(get_style(node.name), ast_text(node.children))
      elsif node.name == 'p'
        @builder.new_layer
        @builder.emit Paragraph.new(get_style(node.name), ast_text(node.children))
      elsif node.name =~ /h(\d+)/
        @builder.new_layer
        @builder.emit Paragraph.new(get_style(node.name), ast_text(node.children))
      elsif node.name == 'ul'
        @builder.new_layer ilvl: true
        unless @builder.nested?
          @definition = numbering.register(get_style(node.name))
        end
        @builder.push_all(node.children)
      elsif node.name == 'ol'
        @builder.new_layer ilvl: true
        unless @builder.nested?
          @definition = numbering.register(get_style(node.name))
        end
        @builder.push_all(node.children)
      elsif node.name == 'li'
        @builder.new_layer
        @builder.emit ListParagraph.new(@definition.style, ast_text(node.children), @definition.numid, @builder.ilvl)
      elsif node.text?
        # SKIP?
      else
        raise ArgumentError, "Don't know how to handle node: #{node.inspect}"
      end
    end

    def ast_text(nodes, format: TextFormat.default)
      runs = nodes.flat_map do |node|
        if node.text?
          Text.new(node.text, format)
        elsif node.name == 'br'
          Newline.new
        elsif node.name == 'strong' || node.name == 'b'
          ast_text(node.children, format: format.with_bold).nodes
        elsif node.name == 'em' || node.name == 'i'
          ast_text(node.children, format: format.with_italic).nodes
        elsif node.name == 'u'
          ast_text(node.children, format: format.with_underline).nodes
        elsif ['ul', 'ol', 'p', 'div'].include?(node.name)
          @builder.push(node)
          nil
        else
          raise ArgumentError, "Don't know how to handle node: #{node.inspect}"
        end
      end
      Collection.new(runs.compact)
    end
  end
end
