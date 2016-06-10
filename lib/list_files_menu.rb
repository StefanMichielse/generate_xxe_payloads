# encoding: ASCII-8BIT

def usage_verify_file(string_replace)
  if File.file?(@options['file'])
    puts "|+| Using #{@options['file']}"
    string_replace ? find_string(nil, 0) : choose_file(@options['file'])
  else
    puts "|!| #{@options['file']} cannot be found. Verify file exists."
  end
  exit
end

def check_file_exist
  if File.file?(@options['file'])
    puts "|+| #{@options['file']} Loaded\n"
    choose_file(@options['file'])
  else
    puts "|!| #{@options['file']} cannot be found. Set with -f or modify config.json"
    exit
  end
end


def list_files_menu(string_replace)
  if(@options['file'].size > 0)
    usage_verify_file(string_replace)
  else
    puts "|+| Using #{@options['file']}"
    check_file_exist
    find_string(nil, 0) if string_replace
  end
end