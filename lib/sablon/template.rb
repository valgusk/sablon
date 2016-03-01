module Sablon
  class Template
    def initialize(path)
      @path = path
    end

    # Same as +render_to_string+ but writes the processed template to +output_path+.
    def render_to_file(output_path, context, properties = {})
      File.open(output_path, 'wb') do |f|
        f.write render_to_string(context, properties)
      end
    end

    # Process the template. The +context+ hash will be available in the template.
    def render_to_string(context, properties = {})
      render(context, properties).string
    end

    private
    def render(context, properties = {})
      Sablon::Numbering.instance.reset!
      Zip.sort_entries = true # required to process document.xml before numbering.xml

      # parse resources
      resources = {}

      resources_xml = Zip::File.open(@path).get_entry('word/_rels/document.xml.rels').get_input_stream.read
      resources_document = Nokogiri::XML.parse(resources_xml)
      resource_collection = resources_document.search('Relationships').first
      relationships = resource_collection.search('Relationship')
      relationships.each { |r| resources[r['Id'].to_s] = r }
      # parse resources end

      Zip::OutputStream.write_buffer(StringIO.new) do |out|
        opened_zip = Zip::File.open(@path)
        opened_zip.each do |entry|
          entry_name = entry.name

          content = entry.get_input_stream.read
          if entry_name == 'word/document.xml'
            out.put_next_entry(entry_name)
            out.write(process(Processor::Document, content, context, properties, resources))
          elsif entry_name =~ /word\/header\d*\.xml/ || entry_name =~ /word\/footer\d*\.xml/
            out.put_next_entry(entry_name)
            out.write(process(Processor::Document, content, context, {}, resources))
          elsif entry_name == 'word/numbering.xml'
            out.put_next_entry(entry_name)
            out.write(process(Processor::Numbering, content))
            # out.write(Processor.process_rels(Nokogiri::XML(content), resources).to_xml)
          elsif entry_name == 'word/_rels/document.xml.rels'
            false
          else
            out.put_next_entry(entry_name)
            out.write(content)
          end
        end

        resources.each do |id, r|
          next if r.is_a?(Nokogiri::XML::Node)

          xml = %{<Relationship Id="#{ id }" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/auto#{ id.downcase }.png"/>}
          node = Nokogiri::XML::parse(xml).children.first
          resource_collection.add_child(node)

          out.put_next_entry(File.join('word', 'media', "auto#{ id.downcase }.png"))
          out.write(File.read(r.data))
          # resources[id] = node
        end

        out.put_next_entry(File.join('word', '_rels', 'document.xml.rels'))
        out.write(resources_document.to_xml(indent: 0).gsub("\n",""))

        binding.pry
      end
    end

    # process the sablon xml template with the given +context+.
    #
    # IMPORTANT: Open Office does not ignore whitespace around tags.
    # We need to render the xml without indent and whitespace.
    def process(processor, content, *args)
      document = Nokogiri::XML(content)
      processor.process(document, *args).to_xml(indent: 0, save_with: 0)
    end
  end
end
