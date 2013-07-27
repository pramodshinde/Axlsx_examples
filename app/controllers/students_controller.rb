class StudentsController < ApplicationController
  before_filter :load_students
  def index
  end

  def generate
    %x[rake generate:data] if @students.empty?
  end

  private
  def load_students
    @students = Student.all
  end
end
