namespace :generate do
  desc "Generating data"
  task :data => [:students, :update]
  task :students => :environment do 
    40.times{ |i| Student.create!(fname: "I am too long please fit me inside_#{i}", lname: "I am also too long please get me inside_#{i}", marks: rand(100)) }
    60.times{ |i| Student.create!(fname: "student_fname_#{i}", lname: "student_lname_#{i}", marks: rand(100)) }
  end
  task :update => :environment do
    students = Student.all
    students.each do |s|
      g = (s.marks > 70) ? "A" : ((s.marks > 50 ) ? "B" : "C" )
      r = (s.marks > 35) ? "PASS" : "FAIL"
      if r == "FAIL"
        s.update_attributes(percentage: s.marks, grade: "", remark: r)
      else
        s.update_attributes(percentage: s.marks, grade: g, remark: r)
      end
    end
  end
end
