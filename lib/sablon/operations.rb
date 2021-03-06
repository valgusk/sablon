# -*- coding: utf-8 -*-
module Sablon
  module Statement
    class Insertion < Struct.new(:expr, :field, :numbering)
      def evaluate(context)
        if content = expr.evaluate(context)
          content = Sablon::Content.wrap(expr.evaluate(context))
          content.numbering = self.numbering if content.respond_to?(:numbering=)
          field.replace(content)
        else
          field.remove
        end
      end
    end

    class Loop < Struct.new(:list_expr, :iterator_name, :block)
      def evaluate(context)
        value = list_expr.evaluate(context)
        value = value.to_ary if value.respond_to?(:to_ary)
        raise ContextError, "The expression #{list_expr.inspect} should evaluate to an enumerable but was: #{value.inspect}" unless value.is_a?(Enumerable)

        content = value.flat_map do |item|
          iteration_context = context.merge(iterator_name => item)
          block.process(iteration_context)
        end
        block.replace(content.reverse)
      end
    end

    class Call < Struct.new(:call_expr, :block, :tail)
      def evaluate(context)
        value = call_expr.evaluate(context)
        content = value.call(block.body, *parse_arguments(context), &processor(context))
        block.replace(content.reverse)
      end

      def processor(context)
        proc do |xml_node, call_context|
          context = context.merge({ call_context: call_context })
          Processor::Document.process xml_node, context, block.resources, block.numbering
        end
      end

      def parse_arguments(context)
        return [] unless tail && tail.is_a?(String)

        arg_strings = tail.match(/^\(([\s\S]+)\)$/).to_a[1].to_s.split(/,\s*/)
        arg_strings.map do |arg|
          begin
            eval(arg)
          rescue SyntaxError => e
            if arg.is_a?(String) && arg[/^=/]
              Expression.parse(arg.sub(/^=/, '')).evaluate(context)
            else
              raise e
            end
          end
        end
      end
    end

    class Condition < Struct.new(:conditon_expr, :block, :predicate)
      def evaluate(context)
        value = conditon_expr.evaluate(context)
        if truthy?(predicate ? value.public_send(predicate) : value)
          block.replace(block.process(context).reverse)
        else
          block.replace([])
        end
      end

      def truthy?(value)
        case value
        when Array;
          !value.empty?
        else
          !!value
        end
      end
    end

    class Comment < Struct.new(:block)
      def evaluate(context)
        block.replace []
      end
    end
  end

  module Expression
    class Variable < Struct.new(:name)
      def evaluate(context)
        context[name]
      end

      def inspect
        "«#{name}»"
      end
    end

    class LookupOrMethodCall < Struct.new(:receiver_expr, :expression)
      def evaluate(context)
        if receiver = receiver_expr.evaluate(context)
          expression.split(".").inject(receiver) do |local, m|
            case local
            when Hash
              local[m]
            else
              local.public_send m if local.respond_to?(m)
            end
          end
        end
      end

      def inspect
        "«#{receiver_expr.name}.#{expression}»"
      end
    end

    def self.parse(expression)
      if expression.include?(".")
        parts = expression.split(".")
        LookupOrMethodCall.new(Variable.new(parts.shift), parts.join("."))
      else
        Variable.new(expression)
      end
    end
  end
end
