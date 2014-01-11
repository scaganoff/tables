# 
#  Author::    Saul Caganoff  (mailto:scaganoff@gmail.com)
#  Copyright:: Copyright (c) 2010, Saul Caganoff
#  License::   Creative Commons Attribution 3.0 Australia License (http://creativecommons.org/licenses/by/3.0/au/)
# 
require 'win32ole'

module Tables

  class ExcelTableReader < TableReader

    attr_reader :workbook, :table

    def initialize(filename=nil)
      @excel=WIN32OLE.new('Excel.Application')
      open_file(filename) unless filename.nil?
      super()
    end

    def open_file(filename)
      path=get_file_path(filename)
      @excel.Workbooks.Open(path)
      puts "Open workbook '#{path}'" if $DEBUG
      @workbook=@excel.Workbooks.Item(1)
      @worksheets=[]
    end

    def create_file(filename)
      path=get_file_path(filename)
      @workbook=@excel.Workbooks.Add
      @worksheets=[]
      @workbook.SaveAs(path)
    end

    def extract_table(worksheet, options={})
      rtf_columns=options[:rtf_columns]
      rtf_columns ||= []
      progress=options[:progress]
      sheet=get_worksheet(worksheet)
      range=sheet.UsedRange
      ncols=range.Columns.Count
      nrows=range.Rows.Count

      result=[]
      (1..nrows).each do |idx|
        row=range.Rows(idx)
        result << ExcelTableReader.extract_row(row,ncols,rtf_columns)
        if progress and idx.modulo(50)==0 then
          percent=(idx.fdiv(nrows)*100).round(0)
          puts ">> extracting row #{idx} (#{percent}%)"
        end
      end
      Table.new(result,sheet.name)
    end

    def write_table(table, worksheet=0, options={})
      progress=options[:progress]
      sheet=get_worksheet(worksheet)
      idx=0
      nrows=table.row_count
      table.each_row do |row|
        idx+=1
        if progress and idx.modulo(50)==0 then
          percent=(idx.fdiv(nrows)*100).round(0)
          puts ">> writing row #{idx} (#{percent}%)"
        end
        r=sheet.Rows(idx)
        row.each_with_index {|val,jdx| r.Cells(jdx+1).Value=val.to_s }
      end
    end

    def write_column(table, column_name, worksheet=0, options={})
      progress=options[:progress]
      sheet=get_worksheet(worksheet)
      values=table.get_column(column_name)
      column_index=table.colindex[column_name]+1
      nrows=table.row_count
      values.each_with_index do |val,idx|
        if progress and idx.modulo(50)==0 then
          percent=(idx.fdiv(nrows)*100).round(0)
          puts ">> updating row #{idx} (#{percent}%)"
        end
        r=sheet.Rows(idx+1)
        c=r.Cells(column_index).Value=val.to_s
      end
    end


    def table_count
      @workbook.WorkSheets.Count
    end

    def clean
      @tables.each do |table|
        table.remove_blank_rows!(1)
        table.remove_repeat_headers!
        table.demerge!
      end
    end

    def save
      @workbook.save
    end

    def exit
      @excel.quit
    end

  private

    def get_worksheet(worksheet)
      if worksheet.is_a?(Integer) then
        if (worksheet+1) > @workbook.Worksheets.Count then
          sheet=@workbook.Worksheets.Add
        else
          sheet=@workbook.Worksheets.Item(worksheet+1)
        end
      else
        sheet=@workbook.Worksheets.Item(worksheet)
      end
      sheet
    end

    def get_file_path(filename)
      fso=WIN32OLE.new('Scripting.FileSystemObject')
      fso.GetAbsolutePathName(filename)
    end

    def ExcelTableReader.extract_row(excel_row,n,rtf_columns)

      # convert zero-based rtf columns into 1-based for internal loop
      rtf_cols=rtf_columns.map {|idx| idx+1 }

      row=[]
      (1..n).each do |idx|
        if rtf_cols.include?(idx) then
          row << slow_extract_text(excel_row.Cells(idx))
        else
          row << extract_text(excel_row.Cells(idx))
        end
      end
      row
    end

    def ExcelTableReader.extract_text(range)
      string=range.Text
      string.sub("\a"," - ")
    end

    def ExcelTableReader.slow_extract_text(range)
      string=""
      n=range.Characters.Count
      (1..n).each do |idx|
        c=range.Characters(idx,1)
        t=c.Text
        if t=="\a" then
          string+=" - "
        else
          string += t unless c.Font.Strikethrough
        end
      end
      string
    rescue
      string=extract_text(range)
    end

    #def ExcelTableReader.extract_text(range)
    #  text = range.Text[0..-3]
    #  list_text = range.ListFormat.ListString
    #  text.empty? ? list_text : text   # return list_text if text is empty
    #end

  end

end