require 'rails_helper'

RSpec.describe OpenSheet, type: :model do
  before do
    @filename = File.join(Rails.root, 'open_sheet.ods')
  end

  describe 'problem' do
    it 'does not work with tempfile' do
      tfile = Tempfile.new(['tfile', '.ods'])
      data = File.open(@filename).read
      tfile.write data
      tfile.rewind
      expect { result = OpenSheet.new tfile }.to raise_error 'no implicit conversion of nil into String'

    end

    it 'works with regular file' do
      result = OpenSheet.new @filename
      expect(result.values.class).to eq Array
      # pp result.values
    end
  end
end
