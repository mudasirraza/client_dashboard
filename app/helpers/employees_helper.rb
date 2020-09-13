module EmployeesHelper
  def valid_excel_import_types
    Importer::Excel::VALID_TYPES.join(",")
  end
end
