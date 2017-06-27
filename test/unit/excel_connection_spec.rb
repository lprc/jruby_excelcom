# encoding: utf-8

require 'minitest/autorun'
require 'tmpdir'
require_relative '../../lib/jruby_excelcom'

describe 'ExcelConnection' do

  $con ||= begin
    e = ExcelConnection::connect
    e.display_alerts = false
    e
  end
  $wb ||= $con.workbook("#{File.dirname(File.absolute_path(__FILE__))}/../resources/test.xlsx")

  Minitest.after_run {
    $wb.close unless $wb.nil?; $wb = nil
    $con.quit unless $con.nil?; $con = nil
  }

  it '::connect' do
    $con.wont_be_nil
    ExcelConnection::connect(false) do |con|
      con.version.is_a? String
    end
  end

  it '#version' do
    $con.version.is_a? String
  end

  it '#workbook' do
    $wb.wont_be_nil
    $con.workbook("#{File.dirname(File.absolute_path(__FILE__))}/../resources/test2.xlsx") { |wb|
      wb.name.is_a? String
    }
  end

  it '#new_workbook' do
    path = "#{Dir.tmpdir}/newwb.xlsx"
    $con.new_workbook(path) { |wb|
      wb.name.must_equal "newwb.xlsx"
    }
    File.exists?(path).must_equal true
    File.delete path
  end

end
