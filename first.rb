# frozen_string_literal: true

require 'rubyXL'
require 'rubyXL/convenience_methods'
# workbook
workbook = RubyXL::Workbook.new
# workbook - worksheet
worksheet = workbook.worksheets[0]
worksheet.sheet_name = "first sheet"
workbook.add_worksheet("seconod sheet")

row = 0
column = 0
# workbook - worksheet - sheet_data
worksheet.add_cell(0, 0, 'A1')
worksheet.sheet_data[row][column].change_contents("aaa")

# drop down
row += 1
column += 1
worksheet.data_validations = RubyXL::DataValidations.new
contents = %w[apple orange banana]
formula = RubyXL::Formula.new(expression: "\"#{contents.join(',')}\"")
# https://github.com/weshatheleopard/rubyXL/blob/fb661a86dd312bd735bef88a6854951f2f5bed56/lib/rubyXL/objects/data_validation.rb#L7
worksheet.data_validations << RubyXL::DataValidation.new(
  sqref: RubyXL::Reference.new(row, column),
  formula1: formula,
  type: 'list',
  prompt_title: nil,
  prompt: nil,
  show_input_message: true,
)
# list
row += 1
column += 1
contents.each_with_index do |content, i|
  worksheet.add_cell(row, column)
  worksheet.sheet_data[row][column].change_contents("#{i}_#{content}")
  row += 1
end
# font
row += 1
column += 1
worksheet.add_cell(row, column)
worksheet.sheet_data[row][column].change_contents("big font")
worksheet.sheet_data[row][column].change_font_size(36)

workbook.write("./file.xlsx")
