class StudentsController < ApplicationController
  before_filter :load_students
  before_filter :load_workbook, except: [:index, :generate]
  def index
  end

  def generate
    %x[rake generate:data] if @students.empty?
  end

  def export
    case params[:type]
    when "Basic"
      exprot_basic_xlsx
    when "Row&Col"
      exprot_row_col_xlsx
    when "Custom"
      export_custom_xlsx
    when "All apply"
      export_all_together_xlsx
    when "Merge"
      export_merge_xlsx
    end
  end

  def exprot_basic_xlsx
    @wb.add_worksheet(name: "Basic") do |sheet|
      sheet.add_row get_header 
      @students.each do |st|
        sheet.add_row [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark]
      end
    end
    @p.serialize("#{Rails.root}/tmp/basic.xlsx")
    send_file("#{Rails.root}/tmp/basic.xlsx", filename: "Basic.xlsx", type: "application/xlsx")
  end

  def exprot_row_col_xlsx
    @wb.add_worksheet(name: "Row&Col") do |sheet|
      sheet.add_row get_header 
      @students.each do |st|
        sheet.add_row [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark]
      end
      sheet.col_style 4, @center, row_offset: 1
      sheet.row_style 0, @header, col_offset: 1
    end
    @p.serialize("#{Rails.root}/tmp/row_col.xlsx")
    send_file("#{Rails.root}/tmp/row_col.xlsx", filename: "Row_Col.xlsx", type: "application/xlsx")
  end

  def export_custom_xlsx
    @p.use_autowidth = false
    @wb.add_worksheet(name: "Custom") do |sheet|
      sheet.add_row get_header, style: @header
      @students.each do |st|
        if st.fname.length >= 21
          sheet.add_row [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark], style: @data, height: 25 
        else
          sheet.add_row [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark], style: @data 
        end
      end
      sheet.column_widths 20, 20, nil, nil, nil, nil
    end
    @p.serialize("#{Rails.root}/tmp/custom.xlsx")
    send_file("#{Rails.root}/tmp/custom.xlsx", filename: "Custom.xlsx", type: "application/xlsx")
  end

  def export_all_together_xlsx
    @wb.add_worksheet(name: "All") do |sheet|
      sheet.add_row get_header, style: @header
      @students.each do |st|
        if st.fname.length >= 21
          if st.remark == "PASS"
            sheet.add_row [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark], style: @style_pass, height: 25
          else
            sheet.add_row [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark], style: @style_fail, height: 25
          end
        else
          if st.remark == "PASS"
            sheet.add_row [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark], style: @style_pass
          else
            sheet.add_row [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark], style: @style_fail
          end
        end
      end
      sheet.column_widths 20, 20, nil, nil, nil, nil
    end
    @p.serialize("#{Rails.root}/tmp/all.xlsx")
    send_file("#{Rails.root}/tmp/all.xlsx", filename: "All.xlsx", type: "application/xlsx")
  end

  def export_merge_xlsx 
    @wb.add_worksheet(name: "All") do |sheet|
      sheet.add_row ["", "Student Result Detail", "", "", "", ""], style: @heading, height: 30
      sheet.merge_cells("B1:D1")
      sheet.add_row get_header, style: @header
      @students_with_a = Student.where(grade: "A") 
      @students_with_b = Student.where(grade: "B") 
      @students_with_c = Student.where(grade: "C")
      @students_with_f = Student.where(grade: "")
      @students_with_a.each do |st|
        data_array = [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark]
        if st.fname.length >= 21
          sheet.add_row data_array, style: @style_pass, height: 25  
        else
          sheet.add_row data_array, style: @style_pass 
        end
      end
      a = @students_with_a.length
      sheet.add_row ["", "Students With Grade A", "=AVERAGE(C3:C#{a+2})", "=AVERAGE(D3:D#{a+2})", "Total", a], style: @total

      @students_with_b.each do |st|
        data_array = [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark]
        if st.fname.length >= 21
          sheet.add_row data_array, style: @style_pass, height: 25  
        else
          sheet.add_row data_array, style: @style_pass 
        end
      end
      b = @students_with_b.length
      sheet.add_row ["", "Students With Grade B", "=AVERAGE(C#{a+4}:C#{a+b+3})", "=AVERAGE(D#{a+4}:D#{a+b+3})", "Total", b], style: @total

      @students_with_c.each do |st|
        data_array = [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark]
        if st.fname.length >= 21
          sheet.add_row data_array, style: @style_pass, height: 25  
        else
          sheet.add_row data_array, style: @style_pass 
        end
      end
      c = @students_with_c.length
      sheet.add_row ["", "Students With Grade C", "=AVERAGE(C#{a+b+4}:C#{a+b+c+4})", "=AVERAGE(D#{a+b+4}:D#{a+b+c+4})", "Total", c], style: @total

      @students_with_f.each do |st|
        data_array = [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark]
        if st.fname.length >= 21
          sheet.add_row data_array, style: @style_fail, height: 25  
        else
          sheet.add_row data_array, style: @style_fail 
        end
      end
      f = @students_with_f.length
      sheet.add_row ["", "Failed Students", "=AVERAGE(C#{a+b+c+4}:C#{a+b+c+f+4})", "=AVERAGE(D#{a+b+c+4}:D#{a+b+c+f+4})", "Total", f], style: @total

      sheet.column_widths 20, 20, nil, nil, nil, nil
    end
    @p.serialize("#{Rails.root}/tmp/Merge.xlsx")
    send_file("#{Rails.root}/tmp/Merge.xlsx", filename: "Merge.xlsx", type: "application/xlsx")
  end

  def export_image_xlsx
    @wb.add_worksheet(name: "Image") do
    end
  end

  private
  def load_students
    @students = Student.all
  end

  def load_workbook
    @p = Axlsx::Package.new
    @wb = @p.workbook
    load_styles
  end

  def load_styles
    @wb.styles do |s| 
      @heading = s.add_style alignment: {horizontal: :center}, b: true, sz: 18, bg_color: "0066CC", fg_color: "FF"
      @header = s.add_style alignment: {horizontal: :center}, b: true, sz: 10, bg_color: "C0C0C0"
      @data = s.add_style alignment: {wrap_text: true}
      @center = s.add_style alignment: {horizontal: :center}, fg_color: "0000FF"
      @green = s.add_style alignment: {horizontal: :left}, fg_color: "00FF00"
      @red = s.add_style alignment: {horizontal: :left}, fg_color: "FF0000"
      @style_pass = [@data, @data, @data, @data, @center, @green]
      @style_fail = [@data, @data, @data, @data, @center, @red]
      @total = [@data, @header, @header, @header, @header, @header]
    end
  end

  def get_header
    ["First Name", "Last Name", "Marks", "Percentage", "Grade", "Remark"]
  end
end
