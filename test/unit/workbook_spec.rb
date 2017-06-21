# encoding: utf-8

require 'minitest/autorun'
require_relative '../../lib/jruby_excelcom'

describe 'Workbook' do

  Minitest.after_run {
    $wb.close unless $wb.nil?; $wb = nil
    $con.quit unless $con.nil?; $con = nil
    File.delete($temp_file_path) if not $temp_file_path.nil? and File.exists?($temp_file_path)
  }

  $con ||= begin
    e = ExcelConnection::connect
    e.display_alerts = false
    e
  end
  $wb ||= $con.workbook("#{File.dirname(File.absolute_path(__FILE__))}/../resources/test.xlsx")
  $temp_file_path ||= "#{Dir::tmpdir}/test.xlsx"

  it '#name' do
    $wb.name.must_equal 'test.xlsx'
  end

  it '#save_as' do
    $wb.save_as($temp_file_path)
    File.exist?($temp_file_path).must_equal true
  end

  it '#worksheet' do
    $wb.worksheet('Tabelle1').wont_be_nil
  end

  it '#add_worksheet' do
    $wb.add_worksheet('test123')
    $wb.worksheet('test123').wont_be_nil
    $wb.worksheet('test123').delete
  end

end