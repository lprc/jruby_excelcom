# encoding: utf-8

require 'java'
require 'jars/jna-4.4.0.jar'
require 'jars/jna-platform-4.4.0.jar'
require 'jars/excelcom-0.0.7.jar'
require 'jruby_excelcom/excel_connection'
require 'jruby_excelcom/workbook'
require 'jruby_excelcom/worksheet'

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
