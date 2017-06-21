# encoding: utf-8

class Workbook

  def initialize(java_wb)
    @wb = java_wb
  end

  # gets the name of this workbook
  def name
    @wb.getName
  end
  alias :getName :name

  # closes this workbook
  # +save+:: whether changes should be saved or not, default is +false+
  def close(save = false)
    @wb.close(save)
  end

  # saves this workbook
  def save
    @wb.save
  end

  # saves this workbook to a new location. Every further operations on this workbook will happen to the newly saved file.
  def save_as(path)
    @wb.saveAs(java.io.File.new(path))
  end
  alias :saveAs :save_as

  # adds and returns a worksheet to this workbook
  # +name+:: name of new worksheet
  def add_worksheet(name)
    Worksheet.new(@wb.addWorksheet(name))
  end
  alias :addWorksheet :add_worksheet

  # gets a worksheet
  # +name+:: name of worksheet to get
  def worksheet(name)
    Worksheet.new(@wb.getWorksheet(name))
  end
  alias :getWorksheet :worksheet

end
