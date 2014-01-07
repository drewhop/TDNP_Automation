# ============================================================================= #
# SuperMeta.rb																	#
# <andrew.weidner@unt.edu>														#
# 																				#
# Description:  creates a super_metadata.xml file based on the information		#
#				recorded in a TDNP microfilm evaluation spreadsheet				#
#				optional export to the PTH New Record Creator					#
#																				#
# INSTRUCTIONS																	#
# 1. cd to the directory with the evaluation spreadsheet						#
#																				#
# 2. run the following code on the command line:								#
#		ruby path\to\SuperMeta.rb filename.xls [optional 2 or 3 flag]			#
#																				#
# 3. script creates a super_metadata.xml file in the working directory			#
#	  * Super-Metadata field 1 in the evaluation spreadsheet is the default		#
#	  * type 2 or 3 after the filename to create metadata for those fields		#
#																				#
# 4. script prompts to load the template in the PTH New Record Creator			#
#	  * requires Chrome browser & ChromeDriver placed on system path			#
#	  * requires selenium-webdriver gem											#
# ============================================================================= #

require 'win32ole'
require 'selenium-webdriver'
require 'io/console'

def create_supermetadata_file

	# create new Excel object
	excel = WIN32OLE.new('Excel.Application')
	puts "Excel failed to start." unless excel
	excel.visible = false # hidden

	# file name is first command line argument
	spreadsheet = File.absolute_path "#{ARGV[0]}"
	workbook = excel.Workbooks.Open("#{spreadsheet}")
	worksheet = workbook.Worksheets(1) # 'Reel Info'

	# optional argument for super-metadata fields 2 & 3
	if (ARGV[1] == '2')
		num = 2
		printed_title = worksheet.Range('c31').value.strip
		serial_title = worksheet.Range('c32').value.strip
		city = worksheet.Range('c33').value.strip
		county = worksheet.Range('f33').value.strip
		lccn = worksheet.Range('c34').value.gsub(/\s+/, "")
		oclc = worksheet.Range('c35').value.strip
		frequency = worksheet.Range('c36').value.strip
		era = worksheet.Range('c37').value.strip
		height = worksheet.Range('c38').value.to_i
		width = worksheet.Range('c39').value.to_i
		name = worksheet.Range('c40').value.strip
		date = worksheet.Range('f40').value.strip

	elsif (ARGV[1] == '3')
		num = 3
		printed_title = worksheet.Range('c44').value.strip
		serial_title = worksheet.Range('c45').value.strip
		city = worksheet.Range('c46').value.strip
		county = worksheet.Range('f46').value.strip
		lccn = worksheet.Range('c47').value.gsub(/\s+/, "")
		oclc = worksheet.Range('c48').value.strip
		frequency = worksheet.Range('c49').value
		era = worksheet.Range('c50').value.strip
		height = worksheet.Range('c51').value.to_i
		width = worksheet.Range('c52').value.to_i
		name = worksheet.Range('c53').value.strip
		date = worksheet.Range('f53').value.strip

	# super-metadata field 1 is default
	else
		num = 1
		printed_title = worksheet.Range('c18').value.strip
		serial_title = worksheet.Range('c19').value.strip
		city = worksheet.Range('c20').value.strip
		county = worksheet.Range('f20').value.strip
		lccn = worksheet.Range('c21').value.gsub(/\s+/, "")
		oclc = worksheet.Range('c22').value.strip
		frequency = worksheet.Range('c23').value.strip
		era = worksheet.Range('c24').value.strip
		height = worksheet.Range('c25').value.to_i
		width = worksheet.Range('c26').value.to_i
		name = worksheet.Range('c27').value.strip
		date = worksheet.Range('f27').value.strip

	end

	filename = File.basename("#{spreadsheet}", ".xls") # grab short file name
	File.open("#{filename}(#{num})_super_metadata.xml", "w") do |supermetadata| # write file

		supermetadata << "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" << "\n"
		supermetadata << "<metadata>" << "\n"
		supermetadata << "  <title qualifier=\"officialtitle\">" << printed_title << " (" << city << ", Tex.)</title>" << "\n"
		supermetadata << "  <title qualifier=\"serialtitle\">" << serial_title << "</title>" << "\n"
		supermetadata << "  <creator>" << "\n"
		supermetadata << "    <type></type>" << "\n"
		supermetadata << "    <name></name>" << "\n"
		supermetadata << "  </creator>" << "\n"
		supermetadata << "  <contributor>" << "\n"
		supermetadata << "    <info></info>" << "\n"
		supermetadata << "    <type></type>" << "\n"
		supermetadata << "    <name></name>" << "\n"
		supermetadata << "  </contributor>" << "\n"
		supermetadata << "  <publisher>" << "\n"
		supermetadata << "    <location></location>" << "\n"
		supermetadata << "    <name></name>" << "\n"
		supermetadata << "  </publisher>" << "\n"
		supermetadata << "  <date qualifier=\"\"></date>" << "\n"
		supermetadata << "  <language>eng</language>" << "\n"
		supermetadata << "  <description qualifier=\"content\">" << frequency << " newspaper from " << city << ", Texas that includes local, state and national news along with advertising.</description>" << "\n"
		supermetadata << "  <description qualifier=\"physical\">pages : ill. ; page " << height << " x " << width << " in.&#13;Digitized from 35 mm. microfilm.</description>" << "\n"
		supermetadata << "  <subject qualifier=\"UNTL-BS\">Business, Economics and Finance - Communications - Newspapers</subject>" << "\n"
		supermetadata << "  <subject qualifier=\"UNTL-BS\">Business, Economics and Finance - Journalism</subject>" << "\n"
		supermetadata << "  <subject qualifier=\"UNTL-BS\">Business, Economics and Finance - Advertising</subject>" << "\n"
		supermetadata << "  <subject qualifier=\"UNTL-BS\">Places - United States - Texas - " << county << " County - " << city << "</subject>" << "\n"
		supermetadata << "  <subject qualifier=\"LCSH\">" << county << " County (Tex.) -- Newspapers.</subject>" << "\n"
		supermetadata << "  <subject qualifier=\"LCSH\">" << city << " (Tex.) -- Newspapers.</subject>" << "\n"
		supermetadata << "  <subject qualifier=\"LCSH\">" << city << " (Tex.) -- Periodicals.</subject>" << "\n"
		supermetadata << "  <primarySource>1</primarySource>" << "\n"
		supermetadata << "  <coverage qualifier=\"placeName\">United States - Texas - " << county << " County - " << city << "</coverage>" << "\n"
		supermetadata << "  <coverage qualifier=\"timePeriod\">" << era << "</coverage>" << "\n"
		supermetadata << "  <source qualifier=\"\"></source>" << "\n"
		supermetadata << "  <citation qualifier=\"\"></citation>" << "\n"
		supermetadata << "  <relation qualifier=\"\"></relation>" << "\n"
		supermetadata << "  <collection>TDNP</collection>" << "\n"
		supermetadata << "  <institution></institution>" << "\n"
		supermetadata << "  <rights qualifier=\"\"></rights>" << "\n"
		supermetadata << "  <resourceType>text_newspaper</resourceType>" << "\n"
		supermetadata << "  <format>text</format>" << "\n"
		supermetadata << "  <identifier qualifier=\"LCCN\">"<< lccn.to_s << "</identifier>" << "\n"
		supermetadata << "  <identifier qualifier=\"OCLC\">" << oclc.to_s << "</identifier>" << "\n"
		supermetadata << "  <degree qualifier=\"\"></degree>" << "\n"
		supermetadata << "  <note qualifier=\"nonDisplay\">Descriptive metadata template by " << name << ": " << date << ".</note>" << "\n"
		supermetadata << "  <meta qualifier=\"hidden\">False</meta>" << "\n"
		supermetadata << "</metadata>"

	end

	t = Time.now
	
	# add export message to spreadsheet
	if (ARGV[1] == '2')
		worksheet.Range('c29').value = "EXPORTED: #{t}"
	elsif (ARGV[1] == '3')
		worksheet.Range('c42').value = "EXPORTED: #{t}"
	else
		worksheet.Range('c16').value = "EXPORTED: #{t}"
	end
	
	workbook.Close(1) # save and close
	excel.Quit
	excel = nil # destroy Excel object

	print "\nFILE EXPORT COMPLETE. Open in New Record Creator? (y/n) "
	nrc = nil
	
	until ((nrc == 'y') || (nrc == 'n')) do
	
		nrc = STDIN.gets.chomp
		
		if (nrc == 'n')
			
			puts "\nGoodbye!"
				
			exit 0
			
		elsif (nrc == 'y')

			data = File.read("#{filename}(#{num})_super_metadata.xml") # get supermetadata
		
			print "\nEnter your NRC username: "
			uname = STDIN.gets.chomp
			print "Enter your NRC password: "
			pword = STDIN.noecho(&:gets).chomp # hide password input on terminal
			print "\n\nNew Record Creator >> "
			
			driver = Selenium::WebDriver.for :chrome
			driver.navigate.to "http://edit.texashistory.unt.edu/nrc/import"
			
			element = driver.find_element(:id, 'id_username')
			element.send_keys "#{uname}"
			element = driver.find_element(:id, 'id_password')
			element.send_keys "#{pword}"
			element.submit

			element = driver.find_element(:name, 'text_input')
			element.send_keys "#{data}"
			element.submit
			
			exit 0
			
		else
			puts "\nERROR: unrecognized input."
			puts "Please enter \"y\" or \"n\": "
			next
		end
	end
	
	rescue WIN32OLERuntimeError # error handling
	puts "\nERROR: #{ARGV[0]} is not a valid file name."
	puts "\tCheck the spelling and directory path and try again."
	abort

	rescue NoMethodError # error handling
	puts "\nERROR: One or more cells may be blank in Super-Metadata field #{num}."
	puts "\tCheck the spreadsheet (#{ARGV[0]}) and try again."
	workbook.close
	excel.quit
	excel = nil
	abort
	
end

create_supermetadata_file