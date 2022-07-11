# frozen_string_literal: true

require 'rubyXL'
require 'rubyXL/convenience_methods'
workbook = RubyXL::Workbook.new
worksheet = workbook.worksheets[0]
worksheet.sheet_name = "first sheet"
worksheet.add_cell(0, 0, 'A1')
workbook.write("./file.xlsx")
