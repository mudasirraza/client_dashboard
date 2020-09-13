module Importer
  class Excel

    VALID_TYPES = %w(
      application/vnd.ms-excel
      application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
      application/vnd.ms-excel.sheet.macroEnabled.12
      application/vnd.oasis.opendocument.spreadsheet
    )

    def initialize file
      @file = file
      @total = 0
      @skipped = 0
      @skipped_reasons = {}
      @errors = []
    end

    def perform
      unless valid_file?
        @errors << "The file should be an excel file"
        return
      end

      data = Roo::Spreadsheet.open(@file)
      headers = generate_headers data
      if headers.blank?
        @errors << "The excel file is missing the required columns"
        return
      end

      data.each_with_index(headers) do |hash, i|
        next if i == 0 #skip headers
        if !hash_is_valid? hash
          @skipped += 1
          @skipped_reasons[i] = 'Invalid record'
          next
        end
        
        employee = Employee.new(
          first_name: hash[:first_name],
          last_name: hash[:last_name],
          company_id: hash[:company_id]
        )
        clients = Client.where(id: hash[:clients].to_s.split(",").map(&:to_i)) if hash[:clients].present?
        employee.clients = clients if clients.present?
        if employee.save
          @total += 1
        else
          @skipped += 1
          @skipped_reasons[i] = employee.errors.full_messages.join(",")
        end
      end
    end

    def get_errors
      @errors
    end

    def get_report
      {
        total: @total,
        skipped: @skipped,
        skipped_reasons: @skipped_reasons
      }
    end

    private

    def valid_file?
      VALID_TYPES.include?(@file.content_type)
    end

    def hash_is_valid? hash
      return false if hash[:company_id].blank? || hash[:first_name].blank? || hash[:last_name].blank?
      return false if Company.find_by_id(name: hash[:company_id])
      true
    end

    def generate_headers data
      headers = data.row(1)
      found_headers = {}
      found_headers[:first_name] = /first[ -]name/i if headers.any? { |h| h =~ /first[ -]name/i }
      found_headers[:last_name] = /last[ -]name/i if headers.any? { |h| h =~ /last[ -]name/i }
      found_headers[:company_id] = /company|company[ -]id/i if headers.any? { |h| h =~ /company|company[ -]id/i }
      found_headers[:clients] = /clients/i if headers.any? { |h| h =~ /clients/i }

      if found_headers[:first_name].blank? || found_headers[:last_name].blank? || found_headers[:company_id].blank?
        found_headers = {}
      end

      found_headers
    end
  end
end
