class OpenSheet
  attr_reader :values

  def initialize(file)
    @workbook = Roo::OpenOffice.new(file)

    read_data
  end

  def read_data
    # pp '--------- pretend data ----------------'
    # pp captured_spreadsheet_data
    # pp '---------- end of pretend data ---------------'

    @workbook.default_sheet = @workbook.sheets[0]

    @values = []
    @workbook.sheets.each_index do |si|
      sv = []
      @workbook.default_sheet = @workbook.sheets[si]
      next if @workbook.first_row.blank?

      @workbook.entries.each_with_index do |row, ri|
        row.each_with_index do |val, ci|
          next unless val
          spanning_data = @workbook.spannings.assoc([ri + 1, ci + 1]) if @workbook.spannings
          if spanning_data
            srow = spanning_data[1][:rows]
            scol = spanning_data[1][:columns]
            sv << ([ri + 1, ci + 1, val, srow, scol])
          else
            sv << ([ri + 1, ci + 1, val, nil, nil])
          end
        end
      end
      @values << sv
    end
  end
end
