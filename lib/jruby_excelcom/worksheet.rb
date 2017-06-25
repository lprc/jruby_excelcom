# encoding: utf-8

class Worksheet

  def initialize(java_ws)
    @ws = java_ws
  end

  # sets the name of this worksheet
  def name=(name)
    @ws.setName(name)
  end
  alias :setName :name=

  # gets the name of this worksheet
  def name
    @ws.getName
  end
  alias :getName :name

  # deletes this worksheet
  def delete
    @ws.delete
  end

  # returns the content in range as a matrix, a vector or a single value, depending on +range+'s dimensions
  # +range+:: range with content to get, default value is _UsedRange_
  def content(range = 'UsedRange')
    c = @ws.getContent(range).to_a.each{ |row| row.to_a }
    columns = c.size
    rows = columns > 0 ? c[0].size : 0
    if columns == 1 and rows == 1 # range is one cell
      c[0][0].is_a?(Java::JavaUtil::Date) ? Time.at(c[0][0].getTime/1000) : c[0][0]
    elsif (columns > 1 and rows == 1) or (columns == 1 and rows > 1) # range is one column or row
        c.flatten.map!{|cell| cell.is_a?(Java::JavaUtil::Date) ? Time.at(cell.getTime/1000) : cell }
    else # range is a matrix
      c.map!{|row| row.map!{|cell| cell.is_a?(Java::JavaUtil::Date) ? Time.at(cell.getTime/1000) : cell }}
    end
  end
  alias :getContent :content

  # sets content in a range
  # +range+:: range in worksheet, e.g. 'A1:B3'
  # +content+:: may be a matrix, a vector or a single value. If it's a matrix or vector, its dimensions must be equal to +range+'s dimensions
  def set_content(range, content)
    if content.is_a?(Array)
      if content[0].is_a?(Array) # content is a matrix
        @ws.java_send :setContent, [java.lang.String, java.lang.Object[][]], range, content
      elsif JavaExcelcom::Util::getRangeSize(range)[0] == 1 # content is a row
        @ws.java_send :setContent, [java.lang.String, java.lang.Object[][]], range, [content]
      else # content is a column
        @ws.java_send :setContent, [java.lang.String, java.lang.Object[][]], range, content.map{|cell| [cell] }
      end
    else # content is a single value
      @ws.java_send :setContent, [java.lang.String, java.lang.Object], range, content
    end
  end
  alias :setContent :set_content

  # sets content in a range
  # +hash+:: must contain +:range+ and +:content+, e.g. <tt>{:range => 'A1:A3', :content => [1,2,3]}</tt>. Otherwise an +ArgumentError+ is raised
  def content=(hash)
    raise ArgumentError, 'cannot set content, argument is not a hash' unless hash.is_a? Hash
    raise ArgumentError, 'cannot set content, hash does not contain :range or :content key' if hash[:range].nil? or hash[:content].nil?
    set_content hash[:range], hash[:content]
  end

  # fills cells in range with color
  # +range+:: range to be colorized
  # +color+:: color to be used, must be an ExcelColor, e.g. ExcelColor::RED
  def set_fill_color(range, color)
    @ws.setFillColor(range, color)
  end
  alias :setFillColor :set_fill_color

  # fills cells in range with color
  # +hash+:: must contain +:range+ and +:color+
  def fill_color=(hash)
    raise ArgumentError, 'cannot set fill color, argument is not a hash' unless hash.is_a? Hash
    raise ArgumentError, 'cannot set fill color, hash does not contain :range or :color key' if hash[:range].nil? or hash[:color].nil?
    set_fill_color hash[:range], hash[:color]
  end

  # gets the fill color of cells in range. Throws a +NullpointerException+ if range contains multiple colors
  def fill_color(range)
    @ws.getFillColor(range)
  end
  alias :getFillColor :fill_color

  # sets font color of cells in range
  # +range+:: range to be colorized
  # +color+:: color to be used, must be an ExcelColor, e.g. ExcelColor::RED
  def set_font_color(range, color)
    @ws.setFontColor(range, color)
  end
  alias :setFontColor :set_font_color

  # sets font color of cells in range
  # +hash+:: must contain +:range+ and +:color+
  def font_color=(hash)
    raise ArgumentError, 'cannot set font color, argument is not a hash' unless hash.is_a? Hash
    raise ArgumentError, 'cannot set font color, hash does not contain :range or :color key' if hash[:range].nil? or hash[:color].nil?
    set_font_color hash[:range], hash[:color]
  end

  # gets the font color of cells in range. Throws a +NullpointerException+ if range contains multiple colors
  def font_color(range)
    @ws.getFontColor(range)
  end
  alias :getFontColor :font_color

  # sets border color of cells in range
  # +range+:: range to be colorized
  # +color+:: color to be used, must be an ExcelColor, e.g. ExcelColor::RED
  def set_border_color(range, color)
    @ws.setBorderColor(range, color)
  end
  alias :setBorderColor :set_border_color

  # sets border color of cells in range
  # +hash+:: must contain +:range+ and +:color+
  def border_color=(hash)
    raise ArgumentError, 'cannot set border color, argument is not a hash' unless hash.is_a? Hash
    raise ArgumentError, 'cannot set border color, hash does not contain :range or :color key' if hash[:range].nil? or hash[:color].nil?
    set_border_color hash[:range], hash[:color]
  end

  # gets the border color of cells in range. Throws a +NullpointerException+ if range contains multiple colors
  def border_color(range)
    @ws.getBorderColor(range)
  end
  alias :getBorderColor :border_color

  # sets comment of cells in range
  # +range+:: range
  # +comment+:: comment text
  def set_comment(range, comment)
    @ws.setComment(range, comment)
  end
  alias :setComment :set_comment

  # sets comment of cells in range
  # +hash+:: must contain +:range+ and +:comment+
  def comment=(hash)
    raise ArgumentError, 'cannot set border color, argument is not a hash' unless hash.is_a? Hash
    raise ArgumentError, 'cannot set border color, hash does not contain :range or :comment key' if hash[:range].nil? or hash[:comment].nil?
    set_comment hash[:range], hash[:comment]
  end

  # gets the comment of cells in range. Throws a +NullpointerException+ if range contains multiple comments
  def comment(range)
    @ws.getComment(range)
  end
  alias :getComment :comment

end
