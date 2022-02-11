require 'nokogiri'
require 'date'
require 'date_core'
require 'zip'

MANIFEST = "imsmanifest.xml"
OUTPUT_XML = "EssayExam.xml"

#v2 - 21 Nov 2018 - Added one more cmd line argument: password (must be single word)
#puts "Expected input: CourseCode Term Midterm/Final Date Time Duration Password"
#puts "Example:  2017-182LGST001G4 '2017-18 Term 2' Midterm 10/3/2019 '10:00:00 AM' '1 hour' coffee"
# ARGV[0] = Course Code
# ARGV[1] = Semester
# ARGV[2] = Midterm/Final
# ARGV[3] = Date (dd/mm/yyyy format)
# ARGV[4] = Time
# ARGV[5] = Duration in hours
# ARGV[6] = Password

class Quiz
	
	def initialize(fileName = "essay_base.xml")
		@template = File.open(fileName) { |f| Nokogiri::XML(f) }
	end
	
	def setName(name)
		@assessment = @template.at_css "assessment"
		@assessment["title"] = name
	end
	
	# duration is given by RO as [x] hrs. Usually, it should be 2 or 3 hrs.
	# for one off usage, this package has to understand the concept of minutes
	# the format accepted is [x] hrs [y] mins
	def setTiming(date, startTime, duration)
		examDateTime = DateTime.strptime(date + " " + startTime,'%d/%m/%Y %I:%M:%S %p')
		
		examDuration = durationInMinutes(duration)
		
		examStartTime = examDateTime - Rational(10, 24 * 60) 
		@assessment.at_css("assess_procextension d2l_2p0|date_start d2l_2p0|month").content = examStartTime.month
		@assessment.at_css("assess_procextension d2l_2p0|date_start d2l_2p0|day").content = examStartTime.mday
		@assessment.at_css("assess_procextension d2l_2p0|date_start d2l_2p0|year").content = examStartTime.year
		@assessment.at_css("assess_procextension d2l_2p0|date_start d2l_2p0|hour").content = examStartTime.hour
		@assessment.at_css("assess_procextension d2l_2p0|date_start d2l_2p0|minutes").content = examStartTime.min
		@assessment.at_css("assess_procextension d2l_2p0|date_start d2l_2p0|seconds").content = 0
		
		examEndTime = examDateTime + Rational(examDuration + 30, 24 * 60)
		@assessment.at_css("assess_procextension d2l_2p0|date_end d2l_2p0|month").content = examEndTime.month
		@assessment.at_css("assess_procextension d2l_2p0|date_end d2l_2p0|day").content = examEndTime.mday
		@assessment.at_css("assess_procextension d2l_2p0|date_end d2l_2p0|year").content = examEndTime.year
		@assessment.at_css("assess_procextension d2l_2p0|date_end d2l_2p0|hour").content = examEndTime.hour
		@assessment.at_css("assess_procextension d2l_2p0|date_end d2l_2p0|minutes").content = examEndTime.min
		@assessment.at_css("assess_procextension d2l_2p0|date_end d2l_2p0|seconds").content = 0
		
		@assessment.at_css("assess_procextension d2l_2p0|time_limit").content = examDuration
		@assessment.at_css("assess_procextension d2l_2p0|grace_period").content = (examDuration/60 > 2)? 3:2
		
		return autosaveEntry(examStartTime, examEndTime)
	end
	
	# returns the day, start/end date in a form accepted by autosave
	# yyyy-mm-dd hh:mm, yyyy-mm-dd hh:mm 
	def autosaveEntry(examStartTime, examEndTime)
		year = examStartTime.year.to_s
		month = format('%02d', examStartTime.month)
		day = format('%02d', examStartTime.mday)
		
		examDate = year+"-"+month+"-"+day 
		
		startTime = format('%02d', examStartTime.hour) + ":" + format('%02d', examStartTime.min)
		endTime = format('%02d', examEndTime.hour) + ":" + format('%02d', examEndTime.min)
		puts(examDate + " " + startTime + "," + examDate + " " + endTime)
		return examDate + " " + startTime + "," + examDate + " " + endTime
	end
	
	# [x] hrs [y] mins return (x*60 + y)
	def durationInMinutes(duration)
		i = duration.delete("^0-9\s").split
		i.length==1 ? i[0].to_i * 60 : i[0].to_i * 60 + i[1].to_i
	end


	def setQuestionText(courseCode, purpose)
		examInfo = courseCode + ' - ' + purpose + ' Exam'
		
# 14 June 2019
# Remove the embedding of Javascript inside the quiz question
# Using D2L autosave going forward
		#stdTemplate = '<p><script type="text/javascript" src="https://eLearn.smu.edu.sg/shared/AutoSave/AutoSave.js">
		#			// <![CDATA[// ]]></script></p><p>'
		#qText = stdTemplate + examInfo + '</p>'
		
		#@assessment.at_css("section mattext").content = qText
		@assessment.at_css("section mattext").content = examInfo
		
		@assessment.at_css("section item")["title"] = "Essay - " + purpose + " Exam"
	end
	
# added 21 Nov 2018
# if this function is not called because there was no call to set the password,
# password will default to 'charity'
	def setPassword(pw)
		@assessment.at_css("assess_procextension d2l_2p0|password").content = pw
	end
	
	def setSubmissionView(semester, purpose)
		@assessment.at_css("assessfeedback mattext").content = '<p><strong><span style="font-size: 18pt; color: #cf2a27;">Your quiz has been submitted successfully.<br /><br />' + 
						semester + " - " + purpose + ' Exam<br /><br />Please show this to the Invigilator before you leave the room.</span></strong></p>'
	end
	
	def output(name = OUTPUT_XML)
		File.write(name, @template)
	end
end

def zipFile(courseCode)
	zipfile_name = courseCode + ".zip"
	input_filenames = [MANIFEST, OUTPUT_XML]
	Zip::File.open(zipfile_name, Zip::File::CREATE) do |zipfile|
		input_filenames.each do |filename|
		# Two arguments:
		# - The name of the file as it will appear in the archive
		# - The original file, including the path to find it
		zipfile.add(filename, filename)
	    end
	end
end

if __FILE__ == $0

q = Quiz.new
q.setName(ARGV[1] + " " + ARGV[2] + " Exam")
q.setTiming(ARGV[3], ARGV[4], ARGV[5])
q.setQuestionText(ARGV[0], ARGV[2])
q.setSubmissionView(ARGV[1], ARGV[2])
#added 21 Nov 2018 - 7th argument - password of quiz
q.setPassword(ARGV[6]) if ARGV.length>6
q.output

zipFile(ARGV[0])

puts "End Program"

end




	