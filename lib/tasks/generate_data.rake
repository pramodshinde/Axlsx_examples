namespace :generate do
  desc "Generating data"
  task :data => [:students, :update]
  task :students => :environment do 
    100.times{ |i| Student.create!(fname: "pramod_#{i}", lname: "shinde_#{i}", marks: rand(100)) }
  end
  task :update => :environment do
    students = Student.all
    students.each do |s|
      g = (s.marks > 70) ? "I" : ((s.marks > 50 ) ? "II" : "III" )
      r = (s.marks > 35) ? "PASS" : "FAIL"
      s.update_attributes(percentage: s.marks, grade: g, remark: r)
    end
  end
end
