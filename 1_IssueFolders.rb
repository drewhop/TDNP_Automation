# 1_IssueFolders.rb
# <andrew.weidner@unt.edu>
# 
# Description:  creates issue folders for 1_TDNP_template.xls
#
# INSTRUCTIONS
# 1. cd to the directory with the completed eval template (filename.xls), for example:
#	Q:\scanned_from_iarchives\1.ReadyToProcess\0_EvaluationSpreadsheets
# 2. enter the following command:  ruby path\to\1_IssueFolders.rb filename.xls

require 'win32ole'
require 'fileutils'

#*********************************************************************
def create_issues(excel, workbook, worksheet, titlepath, row)

	year = worksheet.Cells(row,1).value

	unless year == nil

		year = worksheet.Cells(row,1).value.to_i.to_s.strip
		
		month = worksheet.Cells(row,2).value.to_i.to_s.strip
		if month.length == 1
			month = "0" + month
		end
		
		day = worksheet.Cells(row,3).value.to_i.to_s.strip
		if day.length == 1
			day = "0" + day
		end
		
		volume = worksheet.Cells(row,4).value.to_i.to_s.strip
		if volume == "0"
			volume = ""
		end

		issue = worksheet.Cells(row,5).value.to_i.to_s.strip
		if issue == "0"
			issue = ""
		end

		note = worksheet.Cells(row,6).value
		unless note == nil
			note = note.to_s.strip
		else note = ""
		end

		issuepath = titlepath + "\\" + year + month + day + "01"
		issuedate = year + month + day
		
		unless File.exists?(issuepath)
			FileUtils.mkdir issuepath # create issue folder
			File.open("#{issuepath}\\metadata.txt", "w") do |metadata| # create metadata.txt
				metadata << "volume: " << volume << "\n"
				metadata << "issue: " << issue << "\n"
				metadata << "note: " << note << "\n"
			end
			
			if month == "00"
				puts "MISSING MONTH: #{issuedate}"
			end
			if day == "00"
				puts "MISSING DAY: #{issuedate}"
			end
		end
		
		create_issues(excel, workbook, worksheet, titlepath, row+1)
	end
end
#*********************************************************************
def get_title(excel, workbook, worksheet)

	title = worksheet.Range('c1').value.strip

rescue NoMethodError # title field empty
	workbook.Close(0)
	excel.Quit
	excel = nil # destroy Excel object
	puts "\nERROR: Title field (C1) is empty."
	puts "\tPlease check the value and try again.\n"
	exit	
end
#*********************************************************************
def get_exportpath(excel, workbook, worksheet)

	exportpath = worksheet.Range('c2').value.strip
	
rescue NoMethodError # reel number field empty
	workbook.Close(0)
	excel.Quit
	excel = nil # destroy Excel object
	puts "\nERROR: Reel Path (C2) is empty."
	puts "\tPlease check the value and try again.\n"
	exit	
end
#*********************************************************************
begin
	# create new Excel object
	excel = WIN32OLE.new('Excel.Application')
	puts "Excel failed to start." unless excel
	excel.visible = false # hidden

	# file name is first command line argument
	spreadsheet = File.absolute_path "#{ARGV[0]}"
	workbook = excel.Workbooks.Open("#{spreadsheet}")
	
	worksheet = workbook.Worksheets(1) # Reel Info

	exportpath = get_exportpath(excel, workbook, worksheet)
	title = get_title(excel, workbook, worksheet)
	titlepath = exportpath + "\\" + title
		
	FileUtils.mkdir titlepath unless File.exists?(titlepath)
	
	puts "\n"
		
	# issue folder creation function
	create_issues(excel, workbook, worksheet, titlepath, 4)

	workbook.Close(0)
	excel.Quit
	excel = nil # destroy Excel object

	puts "\nFOLDER CREATION COMPLETE\n"
	puts "#{exportpath}\n"
	exit
	
rescue WIN32OLERuntimeError # invalid spreadsheet name
	excel.Quit
	excel = nil # destroy Excel object
	puts "\nERROR: #{ARGV[0]} is not a valid file name."
	puts "\tCheck the spelling and directory path and try again.\n"
	exit
end
#*********************************************************************
