# 2_IssueFolders.rb
# <andrew.weidner@unt.edu>
# 
# Description:  creates issue folders for 2_TDNP_template.xls
#
# INSTRUCTIONS
# 1. cd to the directory with the evaluation spreadsheet (filename.xls)
# 2. enter the following command:  ruby path\to\2_IssueFolders.rb filename.xls
# 3. creates issue folders in the directory specified by IssueFoldersPath field
# 4. also creates a reel log file for tracking and a notes file

require 'win32ole'
require 'fileutils'

#*********************************************************************
# Folder Creation Logic
#*********************************************************************
def get_calendars(excel, workbook, exportpath, calsheet)

	worksheet = workbook.Worksheets(calsheet)

	year = get_year(excel, workbook, worksheet, calsheet)
	yearstring = year.to_s

	title = get_title(excel, workbook, worksheet, yearstring, calsheet)

	titlepath = exportpath + "\\" + title
	FileUtils.mkdir titlepath unless File.exists?(titlepath)

	yearpath = titlepath + "\\" + yearstring
	FileUtils.mkdir yearpath unless File.exists?(yearpath)

	# issue folders for current calendar worksheet
	create_year(worksheet, yearpath, yearstring)
	
	 # recursive call for next calendar
	get_calendars(excel, workbook, exportpath, calsheet+1)

rescue WIN32OLERuntimeError # ends loop after last worksheet
	return
rescue NoMethodError # export path empty
	workbook.Close(0)
	excel.Quit
	excel = nil # destroy Excel object
	puts "\nERROR: Path to 3.ToDivideIntoIssues in Reel Info is empty."
	puts "\tPlease check the value and try again.\n"
	exit
rescue # invalid export path
	workbook.Close(0)
	excel.Quit
	excel = nil # destroy Excel object
	puts "\nERROR: Invalid Export Path in Reel Info - #{exportpath}"
	puts "\tCheck the spelling and directory path and try again.\n"
	exit
end
#*********************************************************************
def create_year(worksheet, yearpath, yearstring)

	create_month(worksheet, 4, 2, yearpath, yearstring, "01", 1, 1)   # January
	create_month(worksheet, 19, 2, yearpath, yearstring, "02", 1, 1)  # February
	create_month(worksheet, 34, 2, yearpath, yearstring, "03", 1, 1)  # March
	create_month(worksheet, 49, 2, yearpath, yearstring, "04", 1, 1)  # April
	create_month(worksheet, 4, 17, yearpath, yearstring, "05", 1, 1)  # May
	create_month(worksheet, 19, 17, yearpath, yearstring, "06", 1, 1) # June
	create_month(worksheet, 34, 17, yearpath, yearstring, "07", 1, 1) # July
	create_month(worksheet, 49, 17, yearpath, yearstring, "08", 1, 1) # August
	create_month(worksheet, 4, 32, yearpath, yearstring, "09", 1, 1)  # September
	create_month(worksheet, 19, 32, yearpath, yearstring, "10", 1, 1) # October
	create_month(worksheet, 34, 32, yearpath, yearstring, "11", 1, 1) # November
	create_month(worksheet, 49, 32, yearpath, yearstring, "12", 1, 1) # December

end
#*********************************************************************
def create_month(worksheet, row, col, yearpath, yearstring, month, day, weekcount)
	
	if (weekcount <= 5) # row control condition

		# call folder creation function
		day = create_week(worksheet, row, col, yearpath, yearstring, month, day, 1)
		
		# recursive call for next week
		create_month(worksheet, row+3, col, yearpath, yearstring, month, day, weekcount+1)
	end
end
#*********************************************************************
def create_week(worksheet, row, col, yearpath, yearstring, month, day, daycount)
	
	if (daycount <= 7) # column control condition

		if (day <= 9)
			daystring = "0" + day.to_s
		else
			daystring = day.to_s
		end
			
		if (worksheet.Cells(row,col).value.to_i > 0) # if day cell has recorded pages
			
			issuepath = yearpath + "\\" + yearstring + month + daystring + "01"

			volume = worksheet.Cells(row-1,col).value.to_i.to_s.strip
			if volume == "0"
				volume = ""
			end

			issue = worksheet.Cells(row-1,col+1).value.to_i.to_s.strip
			if issue == "0"
				issue = ""
			end
			
			unless File.exists?(issuepath)
				FileUtils.mkdir issuepath # create issue folder
				File.open("#{issuepath}\\metadata.txt", "w") do |metadata| # create metadata.txt
					metadata << "volume: " << volume << "\n"
					metadata << "issue: " << issue << "\n"
					metadata << "note: " << "\n"
				end
			end
		end

		# recursive call for next day
		create_week(worksheet, row, col+2, yearpath, yearstring, month, day+1, daycount+1)
	else
		return day
	end
end
#*********************************************************************

	
#*********************************************************************
# Data Gathering Functions with Empty Field Checks
#*********************************************************************
def get_year(excel, workbook, worksheet, calsheet)

	year = worksheet.Range('q1').value.to_i
	
	if ((year == 0)||(year.to_s.size > 4)||(year.to_s.size < 4)) # check for valid year
		workbook.Close(0)
		excel.Quit
		excel = nil # destroy Excel object
		puts "\nERROR: Year field for worksheet #{calsheet} may be empty or formatted incorrectly."
		puts "\tPlease check the value and try again.\n"
		exit		
	else
		year
	end
end
#*********************************************************************
def get_title(excel, workbook, worksheet, yearstring, calsheet)

	title = worksheet.Range('b1').value.strip

rescue NoMethodError # title field empty
	workbook.Close(0)
	excel.Quit
	excel = nil # destroy Excel object
	puts "\nERROR: Title field for #{yearstring} (worksheet #{calsheet}) is empty."
	puts "\tPlease check the value and try again.\n"
	exit	
end
#*********************************************************************
def get_exportpath(excel, workbook, worksheet)

	exportpath = worksheet.Range('e14').value.strip
	
rescue NoMethodError # reel number field empty
	workbook.Close(0)
	excel.Quit
	excel = nil # destroy Excel object
	puts "\nERROR: Path to Reel Folder field in Reel Info (E14) is empty."
	puts "\tPlease check the value and try again.\n"
	exit	
end
#*********************************************************************
def get_reelnum(excel, workbook, worksheet)

	reelnum = worksheet.Range('c5').value.to_s.strip

rescue NoMethodError # reel number field empty
	workbook.Close(0)
	excel.Quit
	excel = nil # destroy Excel object
	puts "\nERROR: Reel Number field in Reel Info (C5) is empty."
	puts "\tPlease check the value and try again.\n"
	exit	
end
#*********************************************************************
def get_evaluator(excel, workbook, worksheet)

	evaluator = worksheet.Range('c7').value.to_s.strip

rescue NoMethodError # evaluated by field empty
	workbook.Close(0)
	excel.Quit
	excel = nil # destroy Excel object
	puts "\nERROR: Evaluated By field in Reel Info (C7) is empty."
	puts "\tPlease check the value and try again.\n"
	exit	
end
#*********************************************************************
def get_date(excel, workbook, worksheet)

	date = worksheet.Range('g7').value.to_s.strip

rescue NoMethodError # evaluation date field empty
	workbook.Close(0)
	excel.Quit
	excel = nil # destroy Excel object
	puts "\nERROR: Evaluation Date field in Reel Info (G7) is empty."
	puts "\tPlease check the value and try again.\n"
	exit	
end
#*********************************************************************


#*********************************************************************
# Log & Notes Files
#*********************************************************************
def create_log(exportpath, reelnum, evaluator, date)

	logpath = exportpath + "\\" + reelnum + "-log.txt"

	unless File.exists?(logpath)
		File.open("#{logpath}", "w") do |logfile|
			logfile << "Evaluated by: #{evaluator}" << "\n"
			logfile << "Date: #{date}" << "\n"
			logfile << "______________________________" << "\n"
			logfile << "\n"
			logfile << "Separated by: " << "\n"
			logfile << "Date: " << "\n"
			logfile << "______________________________" << "\n"
			logfile << "\n"
			logfile << "Student QC by: " << "\n"
			logfile << "Date: " << "\n"
			logfile << "______________________________" << "\n"
			logfile << "\n"
			logfile << "Staff QC by: " << "\n"
			logfile << "Date: " << "\n"
			logfile << "______________________________" << "\n"
			logfile << "\n"
		end
	else
		puts "The log file for reel #{reelnum} already exists."
	end
end
#*********************************************************************
def create_notes(exportpath, reelnum, worksheet)

	notespath = exportpath + "\\" + reelnum + "-notes.txt"

	unless File.exists?(notespath)
		File.open("#{notespath}", "w") do |notesfile|
			notesfile << "DATE" << "\t\t" << "DUPLICATES" << "\t" << "MISSING" << "\t\t" << "NOTES" << "\n"
			notesfile << "_____________________________________________________________________________\n\n"
		end
		add_note(notespath, worksheet, 2)
	else
		puts "The notes file for reel #{reelnum} already exists."
	end
	
end
#*********************************************************************
def add_note(notespath, worksheet, row)

	date = worksheet.Cells(row,1).value
	
	unless date == nil

		date = date.to_s.strip + "\t"

		duplicates = worksheet.Cells(row,2).value
		unless duplicates == nil
			duplicates = duplicates.to_s.strip
			if duplicates.length > 8
				duplicates = duplicates + "\t"
			else
				duplicates = duplicates + "\t\t"		
			end
		else duplicates = "\t\t"
		end		

		missing = worksheet.Cells(row,3).value
		unless missing == nil
			missing = missing.to_s.strip
			if missing.length > 8
				missing = missing + "\t"
			else
				missing = missing + "\t\t"		
			end
		else missing = "\t\t"
		end

		note = worksheet.Cells(row,4).value
		unless note == nil
			note = note.to_s.strip
		else note = ""
		end
		
		File.open("#{notespath}", "a") do |notesfile| # append note
			notesfile << date << duplicates << missing << note << "\n"
			notesfile << "_____________________________________________________________________________\n\n"
		end
		
		# recursive call for next note
		add_note(notespath, worksheet, row+1)
	end
end
#*********************************************************************


#*********************************************************************
# Flow Control
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

	# log variables (app exits on empty field)
	exportpath = get_exportpath(excel, workbook, worksheet)
	reelnum = get_reelnum(excel, workbook, worksheet)
	evaluator = get_evaluator(excel, workbook, worksheet)
	date = get_date(excel, workbook, worksheet)

	# start issue folder creation
	get_calendars(excel, workbook, exportpath, 4) # <-- firstcalendarposition

	# create log & notes files on successful run
	create_log(exportpath, reelnum, evaluator, date)
	worksheet = workbook.Worksheets(3) # Notes
	create_notes(exportpath, reelnum, worksheet)
	
	workbook.Close(0)
	excel.Quit
	excel = nil # destroy Excel object

	puts "\nFOLDER CREATION COMPLETE"
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
