# frozen_string_literal: true

require 'rubyXL'
require 'rubyXL/convenience_methods'
workbook = RubyXL::Workbook.new
worksheet = workbook[0]
worksheet.sheet_data[0] # Returns first row of the worksheet
worksheet[0]            # Returns first row of the worksheet
workbook.write("./file.xlsx")
