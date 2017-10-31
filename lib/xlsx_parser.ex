require Logger

defmodule XlsxParser do

  alias XlsxParser.XlsxUtil
  alias XlsxParser.XmlParser


  @doc """
  Given a path do an .xlsx and the sheet number (1 based index), this function returns a list of values in the
  sheet. The values are returned as a list of {column, row, value} tuples. An optional parameter of the zip
  processing module is allowed (for testing purposes).
  """
  @spec get_sheet_content(String.t, integer, keyword) :: {:ok, XmlParser.col_row_val} | {:error, String.t}
  def get_sheet_content(path, sheet_number, opts \\ []) do
    skip_validate? =
      Keyword.get(opts, :skip_validate, false)

    zip_module =
      Keyword.get(opts, :zip, :zip)

    should_parse? =
      skip_validate? or match?({:ok, _}, XlsxUtil.validate_path(path))

    with true <- should_parse?,
      {:ok, shared_strings} <- XlsxUtil.get_shared_strings(path, zip_module),
                               Logger.debug("Retrieved shared strings for #{Path.rootname(path)}"),

      {:ok, content}        <- XlsxUtil.get_raw_content(path, "xl/worksheets/sheet#{sheet_number}.xml", zip_module),
                               Logger.debug("Retrieved content for #{Path.rootname(path)}"),

      parsed_xml            <- XmlParser.parse_xml_content(content, shared_strings) do

      Logger.debug("Parsed xml for #{Path.rootname(path)}")

      {:ok, parsed_xml}
    else
      false ->
        XlsxUtil.validate_path(path)

      err ->
        err
    end

  end

  @doc """
  Given a path to an .xlsx, a sheet number, and a path to a csv, this function writes the content of the specified
  sheet to the specified csv path.
  """
  @spec write_sheet_content_to_csv(String.t, integer, String.t, keyword) :: {:ok, String.t} | {:error, String.t}
  def write_sheet_content_to_csv(xlsx_path, sheet_number, csv_path, opts \\ []) do
    file_module =
      Keyword.get(opts, :file, File)

    case get_sheet_content(xlsx_path, sheet_number, opts) do
      {:error, reason} -> {:error, reason}
      {:ok, content}   ->
        csv = XlsxUtil.col_row_vals_to_csv(content)
        case file_module.write(csv_path, csv) do
          {:error, reason} -> {:error, "Error writing csv file: #{inspect reason}"}
          :ok              -> {:ok, csv}
        end
    end
  end

@doc """
Given a path to an .xlsx, this function returns an array of worksheet names
"""
  @spec get_worksheet_names(String.t, keyword) :: {:ok, [String.t]} | {:error, String.t}
  def get_worksheet_names(path, opts \\ []) do
    zip_module =
      Keyword.get(opts, :zip, :zip)

    case XlsxParser.XlsxUtil.get_raw_content(path, "xl/workbook.xml", zip_module) do
      {:error, reason} -> {:error, reason}
      {:ok, content} ->
        import SweetXml
        {:ok, xpath(content, ~x"//workbook/sheets/sheet/@name"l)
                |> Enum.map(&List.to_string(&1))}
    end
  end
end
