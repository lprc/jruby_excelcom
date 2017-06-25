# encoding: utf-8

require 'java'
require_relative 'jars/jna-4.4.0.jar'
require_relative 'jars/jna-platform-4.4.0.jar'
require_relative 'jars/excelcom-0.0.6.jar'
require_relative 'jruby_excelcom/excel_connection'
require_relative 'jruby_excelcom/workbook'
require_relative 'jruby_excelcom/worksheet'

java_import 'excelcom.api.ExcelException'

module JavaExcelcom
  java_import 'excelcom.api.ExcelConnection'
  java_import 'excelcom.util.Util'
  java_import 'excelcom.api.ExcelColor'
end

module ExcelColor
  # derive constants from java enum ExcelColor
  JavaExcelcom::ExcelColor.constants.each{|c| ExcelColor::const_set(c, JavaExcelcom::ExcelColor::const_get(c)) }
end
