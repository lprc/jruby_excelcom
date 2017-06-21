# jruby_excelcom

jruby_excelcom is a wrapper for the java excel modification library <b>excelcom</b> and thus only available for JRuby. 
It uses JNA and MS COM interface to create excel instances and modify excel files.

## Requirements
- MS Windows OS
- MS Office installed (at least Excel)
- Java 1.6 or higher

## How to get
    jruby -S gem install jruby_excelcom

## How to use

    require 'jruby_excelcom'
    
     
    # if connected with block, COM will be uninitialized automatically
    ExcelConnection::connect{ |con|
      wb = con.workbook "path/to/wb.xlsx"
      ws = wb.worksheet 'sheet1'
      
      puts ws.content #=> content in used range
      puts ws.content('A1:B3') #=> content in range A1 to B3
      
      # set content using an assignment with a hash, or by calling set_content
      ws.content = { :range => 'A2:B3', :content => [[1,2],['abc','äöü']] }
      ws.set_content('A4:A6', [2,3,5])
      
      # get some color in there
      ws.set_fill_color('A1', ExcelColor::RED)
      ws.set_font_color('A1', ExcelColor::YELLOW)
      ws.set_border_color('A1', ExcelColor::PINK)
      
      ws_new = wb.add_worksheet 'newsheet' # create a new sheet ...
      ws_new.delete # ... and delete it again
    }
     
    # connect to an existing instance without block, quit must be called, 
    # otherwise COM will be left uninitialized!
    con = ExcelConnection::connect(true)
    # ...
    con.quit
