# encoding: utf-8

class ExcelConnection
  # initializes com and connects to an excel instance
  # +use_active_instance+:: whether an existing excel instance should be used or a new isntance should be created. Default value is +false+
  def initialize(use_active_instance = false)
    @con = JavaExcelcom::ExcelConnection::connect(use_active_instance)
  end

  # see +new+
  # optional block possible, where <tt>ExcelConnection#quit</tt> gets called on blocks end
  # e.g. <tt>ExcelConnection::connect{|con| con.workbook ... }</tt>
  def self.connect(use_active_instance = false)
    con = self.new(use_active_instance)
    if block_given?
      yield(con)
      con.quit
    else
      con
    end
  end

  # initializes com manually, not recommended! happens automatically when an instance is created
  def self.initialize_com
    JavaExcelcom::ExcelConnection::initialize_com
  end

  # uninitializes com manually, not recommended! should happen automatically when <tt>ExcelConnection#quit</tt> is called.
  def self.uninitialize_com
    JavaExcelcom::ExcelConnection::uninitialize_com
  end

  # whether the excel instance should be visible or not
  def visible=(v)
    @con.setVisible v
  end
  alias :setVisible :visible=

  # whether dialog boxes should show up or not (e.g. when saving and overwriting a file)
  def display_alerts=(da)
    @con.setDisplayAlerts(da)
  end
  alias :setDisplayAlerts :display_alerts=

  # gets excel version
  def version
    @con.getVersion
  end
  alias :getVersion :version

  # quits the excel instance and uninitializes com
  def quit
    @con.quit
  end

  # gets the active workbook
  def active_workbook
    Workbook.new(@con.getActiveWorkbook)
  end
  alias :getActiveWorkbook :active_workbook

  # opens a workbook. Optional block possible where workbook gets closed on blocks end, e.g.
  # <tt>con.workbook{|wb| puts wb.name }</tt>
  # +file+:: workbook to be opened. Can be a string or a file object
  def workbook(file)
    if file.is_a? String
      wb = Workbook.new(@con.openWorkbook(java.io.File.new(file)))
    else
      wb = Workbook.new(@con.openWorkbook(java.io.File.new(file.path)))
    end
    if block_given?
      yield(wb)
      wb.close
    else
      wb
    end
  end
  alias :openWorkbook :workbook
  alias :open_workbook :workbook

  # creates a new workbook. If block is given, workbook will be saved and closed at end
  def new_workbook(file)
    if file.is_a? String
      wb = Workbook.new(@con.newWorkbook(java.io.File.new(file)))
    else
      wb = Workbook.new(@con.newWorkbook(java.io.File.new(file.path)))
    end
    if block_given?
      yield(wb)
      wb.close true
    else
      wb
    end
  end
  alias :add_workbook :new_workbook

end
