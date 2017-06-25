# encoding: utf-8

require 'minitest/autorun'
require 'time'
require_relative '../../lib/jruby_excelcom'

describe 'Worksheet' do

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

  before do
    $ws = $wb.add_worksheet 'test'
  end

  after do
    $ws.delete
    $ws = nil
  end

  it '#content' do
    range = 'A2:B5'
    content = [['A22', 123], [54.6, 23.5], ['äöüß', '#?-'], [Time::parse('2017-03-03'), 1234567]]
    $ws.set_content(range, content)

    actual = $ws.content(range)
    actual[0][0].must_equal content[0][0]
    actual[0][1].must_equal content[0][1]
    actual[1][0].must_equal content[1][0]
    actual[1][1].must_equal content[1][1]
    actual[2][0].must_equal content[2][0]
    actual[2][1].must_equal content[2][1]
    actual[3][0].must_equal content[3][0]
    actual[3][1].must_equal content[3][1]

    actual = $ws.content(JavaExcelcom::Util.boundsToRange(1, 0, 4, 3))
    actual[0][0].must_equal content[0][0]
    actual[0][1].must_equal content[0][1]
    actual[1][0].must_equal content[1][0]
    actual[1][1].must_equal content[1][1]
    actual[2][0].must_equal content[2][0]
    actual[2][1].must_equal content[2][1]
    actual[3][0].must_equal content[3][0]
    actual[3][1].must_equal content[3][1]
  end

  it '#content=' do
    # check single cell
    range = 'A1'
    content = 123
    $ws.content = {:range => range, :content => content}
    actual = $ws.content range
    actual.must_equal content

    # check column
    range = 'A2:A3'
    content = [123, 456]
    $ws.content = {:range => range, :content => content}
    actual = $ws.content range
    actual.must_equal content

    # check row
    range = 'A4:B4'
    content = [123, 654]
    $ws.content = {:range => range, :content => content}
    actual = $ws.content range
    actual.must_equal content

    # check matrix
    range = 'A5:B6'
    content = [[235, 7911], [13,17]]
    $ws.content = {:range => range, :content => content}
    actual = $ws.content range
    actual.must_equal content
  end

  it '#fill_color' do
    range = 'A1'
    color = ExcelColor::RED
    $ws.fill_color = {:range => range, :color => color}
    actual = $ws.fill_color range
    actual.must_equal color
  end

  it '#font_color' do
    range = 'A1'
    color = ExcelColor::RED
    $ws.font_color = {:range => range, :color => color}
    actual = $ws.font_color range
    actual.must_equal color
  end

  it '#border_color' do
    range = 'A1'
    color = ExcelColor::RED
    $ws.border_color = {:range => range, :color => color}
    actual = $ws.border_color range
    actual.must_equal color
  end

  it '#comment' do
    range = 'A1'
    comment = "test"
    $ws.comment = {:range => range, :comment => comment}
    actual = $ws.comment range
    actual.must_equal comment
  end

end