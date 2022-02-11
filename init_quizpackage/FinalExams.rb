require 'win32ole' #https://msdn.microsoft.com/en-us/library/bb208132(v=office.12).aspx
require 'csv'
require 'set'  #needed to store venue and section information
require 'date_core'
require_relative 'QuizPackage'

#USAGE
# ARGV[0] = csv file name 
# ARGV[1] = Semester e.g "2017-18 Term 2"
# ARGV[2] = Purpose: Midterm/Final

HEADER_EXAM_DATE = "Exam Date"
HEADER_DAY = "Day"
HEADER_START_TIME = "Start Time"
HEADER_DURATION = "Duration"
HEADER_SUBJECT = "Subject"
HEADER_NBR = "Catalog Nbr"
HEADER_SECTION = "Section"
HEADER_INSTRUCTOR = "Instructor Name"
HEADER_VENUE = "Venue"
HEADER_COMMENTS = "Instr Remarks"
HEADER_EMAIL = "Email"

class Instructor
	attr_accessor :name, :email
	attr_reader :sessions, :courses
	
	def initialize
		@sessions = Array.new()
		@courses = Hash.new { |hash,key| hash[key] = CourseExam.new(key) }
	end
	
	def addSession(session)	
		@sessions << session	
		code = session.subject + session.catalog_num
		@courses[code].addSection(session.section)
		@courses[code].addVenue(session.venue)
		@courses[code].day = session.day
		@courses[code].date = session.date
		@courses[code].time = session.time
		@courses[code].duration = session.duration
		@courses[code].addComments(session.comments)
	end
	
	def <<(session)
		addSession(session)
	end
end

class CourseExam
	attr_reader :code, :venues, :comments, :sections
	attr_accessor :day, :date, :time, :duration
	
	def initialize(code)
		@code = code
		@sections = SortedSet.new
		@venues = SortedSet.new
		@comments = Set.new
	end
	
	def addSection(section)
		@sections.add(section)
	end
	
	def addVenue(venue)
		@venues.add(venue)
	end
	
	def addComments(comment)
		@comments.add(comment)
	end
	
	def comments()
		@comments.to_a.join(",")
	end
end

class ExamSession	

	attr_reader :subject, :catalog_num, :section, :date, :time, :duration, :venue, :day
	attr_accessor :comments
	
	def initialize(subject, catalog, section, day, date, time, duration, venue)
		@subject = subject
		# csv file will have '_ before the numbers so 001 will not become 1
		# need to strip out everything non-alphanumeric
		@catalog_num = catalog.gsub(/[^0-9A-Za-z]/i,"")
		@section = section
		@day = day
		@date = date
		@time = time
		@duration = duration
		@venue = venue
	end
end

class Email
	FORMATTING = "style='font-size:11.0pt;font-family:Calibri,sans-serif;color:#44546A'"
	DEFAULT_EMAIL_BODY = "Dear Prof,<br/><br/>We have received the confirmed list from Office of Registrar (RO) that you would like to conduct online Final Exam and requested for eLearn and IT Support for <br/><br/>"

	DEFAULT_SETUP_Q = "<br/><br/>May I know what is your requirement so that I can setup the exam in your course in eLearn or to provide the necessary IT support for you on the day of the exam?"
	
	DEFAULT_SELF_SETUP = "<br/><br/>I will be arranging the necessary support for your exam. I understand that you will be setting up the exam yourself. Do let me know if you require any assistance for that, and if you have any other IT requirements for your final exam. "

	DEFAULT_LAW = "<br/><br/>Can I check if a single textbox with no HTML Editor will be sufficient for your exam? If so, can I know if you will prefer the quiz to be in the individual or the combined sections?"
	
	DEFAULT_OPEN_BOOK_NOTICE = "<br/><br/>I will also like to take this opportunity to remind you to inform your students that for open-book exams, all students are to bring hard copies of their reference material." +
								"No electronic device will be allowed for use as a means to access soft copies of the reference materials."
								
	WITH_MIDTERM = "<br/><br/>Can I confirm that the same setup as per your midterm will be fine with you?"
	
	DEFAULT_PREV = "<br/><br/>Can I confirm that it will be fine to setup your final exam as per previous semesters, a single textbox (no HTML Editor) and students are supposed to number their answers accordingly?"
								
	MAIL_ENDING = "<br/><br/>Hope to hear from you soon. Thank you."
	
	INSTR_REMARKS = "<br/><br/>Instr Remarks :"
	
	SIGNOFF = "<br/><br/>Regards,<br/>Tee Seng<br/><br/>IITS - eLearn Support"
	
	def initialize
		@outlook = WIN32OLE.new("Outlook.Application")
	end
	
	def draft(instructor_name, email, courses)
		message = @outlook.CreateItem(0)
		message.subject = ARGV[1] + " Online Final Exam IT Support Requirements"
		
		info = "<ul " + FORMATTING + ">"
		comments = "INSTRUCTOR_COMMENTS<br/>"
		courses.each do |k,v|
			info = info + "<li>" + k + " " + formatSections(v.sections) + "</li><ul><li>"
			formatDate(v.date)
			info = info + v.day + ", " +  formatDate(v.date) + ", " + v.time + " (" + formatDuration(v.duration) + ")" + "</li>"
			info += ("<li>" + formatVenues(v.venues) + "</li>")
			info += "</ul>"
			comments << v.comments
		end
		info += "</ul>"
		message.HTMLBody = (fontFormatting(DEFAULT_EMAIL_BODY + info) +
							 fontFormatting(DEFAULT_SETUP_Q + DEFAULT_SELF_SETUP + DEFAULT_PREV+ WITH_MIDTERM + DEFAULT_LAW + DEFAULT_OPEN_BOOK_NOTICE + MAIL_ENDING) +
							 fontFormatting(comments) + fontFormatting(SIGNOFF))
		message.Recipients.Add email
		message.SaveAs('D:\Projects\Exams\\' + instructor_name + '\draft.msg', 3)	
	end
	
	#sections are held as a set. 
	def formatSections(sections)
		sections.to_a.join(",")
	end
	
	#converts 29/11/2017 to 29 Nov 2017
	def formatDate(date)
		d = DateTime.strptime(date,'%d/%m/%Y')
		d.strftime('%-d %-b %Y')
	end
	
	# 2 hrs => 2 hours
	def formatDuration(duration)
		duration.gsub("hrs", "hours")
	end
	
	# "SIS Seminar Rm 3.1" | "SIS Seminar Rm 3.2" => SIS Seminar Rm 3.1, 3.2
	# assumes same building for the same course
	def formatVenues(venues)
		v = ""
		venues.each_with_object(v) do |venue, acc|
			if (acc.empty? == true)
				acc << venue
			else
				similarity = (acc.split & venue.split)
				acc << (", " + venue.gsub(similarity.join(" "), ""))
			end
		end
		v
	end
	
	def fontFormatting(blk)
		"<p " + FORMATTING + ">" + blk + "</p>"
	end
end

def createQuizPackage(instructor_name, exam, courseCode)
	quiz = Quiz.new
	quiz.setName(ARGV[1] + " " + ARGV[2] + " Exam")
	quiz_timing = quiz.setTiming(exam.date, exam.time, exam.duration)
	quiz.setQuestionText(courseCode, ARGV[2])
	quiz.setSubmissionView(ARGV[1], ARGV[2])
	quiz.output
	zipFile('D:\Projects\Exams\\' + instructor_name + '\\' +courseCode)
	
	return quiz_timing
end

txt_autosave = File.open("new_autosave.txt", "w")
abort("Unable to create text file for autosave settings") if txt_autosave.nil?


AllExams = Hash.new{ |hash, key| hash[key] = Instructor.new }
#Taken in as 2017-18 Term 2. Our course codes are 2017-182
quizPackagePrefix = ARGV[1].gsub("Term", "").gsub(/\s/,"")

#Parsing the CSV file
CSV.foreach(ARGV[0], :headers => true) do |row|
	session = ExamSession.new(row[HEADER_SUBJECT], row[HEADER_NBR], row[HEADER_SECTION], row[HEADER_DAY], row[HEADER_EXAM_DATE], row[HEADER_START_TIME], row[HEADER_DURATION], row[HEADER_VENUE])
	session.comments = row[HEADER_COMMENTS]
	AllExams[row[HEADER_INSTRUCTOR]].email = row[HEADER_EMAIL]
	AllExams[row[HEADER_INSTRUCTOR]] <<(session)
	#commented out for testing, else too tedious to delete after every run
	Dir.mkdir(row["Instructor Name"]) unless Dir.exist?(row["Instructor Name"])
end

AllExams.each do |name,obj|
	email = Email.new
	email.draft(name, obj.email, obj.courses)
	obj.courses.each do |h,k|
		courseExam = k
		courseCode =  quizPackagePrefix + k.code
		k.sections.each do |section|
			timing = createQuizPackage(name, k, courseCode+section)
			txt_autosave.puts(courseCode+section+","+timing)
		end
		
		if (k.sections.size > 1)
			combined = "G" + k.sections.to_a.join("-").gsub("G","")
			timing = createQuizPackage(name, k, courseCode+combined)
			txt_autosave.puts(courseCode+combined+","+timing)
		end
	end
end

txt_autosave.close 

puts "End Program"