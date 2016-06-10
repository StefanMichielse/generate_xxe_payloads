# encoding: ASCII-8BIT
def add_payload_of(fz,payloadx,of)

  # get file ext
  ext = @options['file'].split('.').last
  nname = "output_#{Time.now.to_i}_all"
  rand_file = "./output/#{nname}.#{ext}"
  FileUtils::copy_file(@options['file'], rand_file)

  fz.each do |name|
    document = ''
    # Read in the XLSX and grab the document.xml
    Zip::Archive.open(rand_file, Zip::CREATE) do |zipfile|
      zipfile.fopen(name) do |f|
        document = f.read # read entry content
      end
    end

    docx_xml = payload(document,payloadx)

    Zip::Archive.open(rand_file, Zip::CREATE) do |zipfile|
      zipfile.add_or_replace_buffer(name, docx_xml)
    end
  end
  puts "|+| Created #{rand_file}"
end