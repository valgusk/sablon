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
        contents = {}

        opened_zip.each do |entry|
          entry_name = entry.name

          content = entry.get_input_stream.read

          if entry_name == 'word/document.xml'
            contents[entry_name] = process(Processor::Document, content, context, resources, properties)
          elsif entry_name =~ /word\/header\d*\.xml/ || entry_name =~ /word\/footer\d*\.xml/
            contents[entry_name] = process(Processor::Document, content, context, resources)
          elsif entry_name == 'word/numbering.xml'
            contents[entry_name] = process(Processor::Numbering, content)
            # out.write(Processor.process_rels(Nokogiri::XML(content), resources).to_xml)
          elsif entry_name == 'word/_rels/document.xml.rels'
            false
          elsif entry_name == '[Content_Types].xml'
            types = Nokogiri::XML.parse(content)
            ['jpeg', 'png', 'bmp'].each do |type|
              parent = types.children.first
              unless parent.children.select{ |t| t['ContentType'] == "image/#{ type }" }.any?
                parent.add_child(%{<Default Extension="#{ type }" ContentType="image/#{ type }"/>})
              end
            end

            contents[entry_name] = types.to_xml(indent: 0).gsub("\n","")
          else
            contents[entry_name] =content
          end
        end

        resources.each do |id, r|
          next if r.is_a?(Nokogiri::XML::Node)

          name = "auto#{ id.downcase }.#{ r.spec.content_type[/png|jpeg|bmp/] }"
          xml = %{<Relationship Id="#{ id }" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/#{ name }"/>}
          node = Nokogiri::XML::parse(xml).children.first
          resource_collection.add_child(node)

          content = case r.data
                    when ::File, ::Tempfile
                      File.read(r.data)
                    when ::StringIO
                      r.data.read
                    end

          contents[File.join('word', 'media', name)] = content
          # resources[id] = node
        end

        contents[File.join('word', '_rels', 'document.xml.rels')] = resources_document.to_xml(indent: 0).gsub("\n","")

        contents.keys.sort.each do |entry_name|
          out.put_next_entry(entry_name)
          out.write(contents[entry_name])
        end
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
