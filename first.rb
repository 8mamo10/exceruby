# frozen_string_literal: true

require 'rubyXL'
require 'rubyXL/convenience_methods'
# workbook
workbook = RubyXL::Workbook.new
# workbook - worksheet
worksheet = workbook.worksheets[0]
worksheet.sheet_name = "first sheet"
worksheet.add_cell(0, 0, 'A1')

# workbook - worksheet - sheet_data
worksheet.sheet_data[0][0].change_contents("aaa")

# drop down
worksheet.data_validations = RubyXL::DataValidations.new
contents = %w[apple orange banana]
formula = RubyXL::Formula.new(expression: "\"#{contents.join(',')}\"")
# https://github.com/weshatheleopard/rubyXL/blob/fb661a86dd312bd735bef88a6854951f2f5bed56/lib/rubyXL/objects/data_validation.rb#L7
worksheet.data_validations << RubyXL::DataValidation.new(
  sqref: RubyXL::Reference.new(1, 1),
  formula1: formula,
  type: 'list',
  prompt_title: nil,
  prompt: nil,
  show_input_message: true,
)

workbook.write("./file.xlsx")
