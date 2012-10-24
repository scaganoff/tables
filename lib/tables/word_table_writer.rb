#
#  Author::    Saul Caganoff  (mailto:scaganoff@gmail.com)
#  Copyright:: Copyright (c) 2010, Saul Caganoff
#  License::   Creative Commons Attribution 3.0 Australia License (http://creativecommons.org/licenses/by/3.0/au/)
#
require 'win32ole'

module Tables

  class WordTableWriter
    attr_reader :doc

    def initialize
      @app=WIN32OLE.new 'Word.Application'
      @doc=@app.Documents.Add
      super()
    end

    def save_as(filename)
      fso=WIN32OLE.new('Scripting.FileSystemObject')
      @path=fso.GetAbsolutePathName(filename)
      @doc.SaveAs(@path)
    end

    def exit(save_changes=true)
      code=-1
      code=0 unless save_changes
      @app.quit(code)
    end

    def append_table(table, caption=nil)
      @app.Selection.EndKey(6)
      word_table=@doc.Tables.Add(@app.Selection.Range, table.row_count, table.column_count)
      table.each_row_with_index do |row, idx|
        word_row=word_table.Rows(idx+1)
        row.each_with_index do |cell_value, jdx|
          begin
            word_row.Cells(jdx+1).Range.Text=cell_value.to_s
          rescue Exception=>e
            $stderr.puts "Error writing cell number '#{jdx+1}'"
            $stderr.puts e.message
            $stderr.puts e.backtrace
          end
        end
      end

      style_borders(word_table)
      set_caption(word_table,caption) unless caption.nil?
    end

  private
    def style_borders(word_table)
      (1..6).each do |border|
        word_table.Borders(border).LineStyle=1
      end

      (1..4).each do |border|
        word_table.Borders(border).LineWidth=12
      end
      word_table.Rows(1).Borders(3).LineWidth=12
    end

    def set_caption(word_table, caption)
      word_table.Range.InsertCaption("Table",": #{caption}",nil,0)
    end

  end

end

