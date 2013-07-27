class Student
  include Mongoid::Document
  field :fname, type: String
  field :lname, type: String
  field :marks, type: Integer
  field :percentage, type: Float
  field :grade, type: String
  field :remark, type: String
end
