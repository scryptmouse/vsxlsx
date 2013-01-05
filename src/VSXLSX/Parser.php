<?php
namespace VSXLSX;

/**
 * Calculate the alphabetical column name
 * of an index, in the style of 'a', 'b' .. 'z', 'aa', 'ab'
 *
 * @param	int
 * @return	string
 */
function alphacize($idx)
{
	for($r = ""; $idx >= 0; $idx = intval($idx / 26) - 1)
		$r = chr($idx%26 + 0x61) . $r;
	return $r;
}

/**
 * Calculate the numeric column index
 * when given an alphabetical index like 'AA'
 *
 * @param	string
 * @return	int
 */
function indicize($alpha)
{
	$alpha = strtolower($alpha);
	$l = strlen($alpha);
	$n = 0;

	for($i = 0; $i < $l; $i++)
		$n = $n*26 + ord($alpha[$i]) - 0x60;

	return $n-1;
}

/**
 * Get the numerical column index based on $idx
 *
 * @param	string|int
 * @return	int|NULL
 */
function column_index($idx)
{
	$calculated = null;

	if (is_numeric($idx)) {
		$calculated = (int) $idx;
	} elseif (is_string($idx)) {
		$calculated = indicize($idx);
	}

	return $calculated;
}

/**
 * Get the alphabetical column name based on $name
 *
 * @param	string|int
 * @return	string
 */
function column_name($name)
{
	return is_numeric($name) ? alphacize((int) $name) : (string) $name;
}

/**
 * Get the appropriate row index from a cell
 *
 * @param	\SimpleXMLElement
 * @return	int
 */
function row_index(\SimpleXMLElement $cell)
{
	$coords = (string) $cell->attributes()->r;
	$row_idx = preg_replace('/\d+$/', '', $coords);
	return indicize($row_idx);
}

/**
 * Generate an array of default header names with desired length.
 *
 * @param	int|array	Length, or array to calculate length based on
 * @return	array
 */
function default_columns($length)
{
	if (is_array($length))
		$length = count($length);

	$names = array();

	for ($i = 0; $i <= $length; $i++)
		$names[$i] = alphacize($i);

	return $names;
}

/**
 * Recursively remove a directory, equivalent to `rm -rf`.
 *
 * @return	boolean
 */
function recursive_rm($filename) {
	$writable = is_writable($filename);

	if ($writable && is_dir($filename)) {
		$can_delete_this_directory = true;
		$files = array_diff(scandir($filename), array('.', '..'));

		foreach($files as $file)
		{
			if (!recursive_rm($filename.'/'.$file))
				$can_delete_this_directory = false;
		}

		return $can_delete_this_directory ? rmdir($filename) : false;
	} elseif ($writable && is_file($filename)) {
		return unlink($filename);
	} else {
		return false;
	}
}

class Parser
{
	/** @var string Path to spreadsheet */
	protected $file;
	/** @var int Selected worksheet number */
	protected $sheetNum;
	/** @var \SimpleXMLElement sharedStrings XML */
	protected $strings;
	/** @var \SimpleXMLElement Worksheet XML */
	protected $sheet;
	/** @var array An array of column names */
	protected $headers = array();
	/** @var array An array that holds parsed rows. */
	protected $parsed = array();
	/** @var array An array of error messages, populated if some step fails */
	protected $errors = array();
	/** @var boolean Whether or not the parsed spreadsheet has a "header" row to derive column names from */
	protected $has_header = true;
	/** @var array An array of header names */
	protected $header_overrides = array();
	/** @var string Directory to store temporary files */
	protected $base_tmp_dir = './tmp';

	/**
	 * @param	string	Path to XLSX file.
	 * @param	int	Selected worksheet number
	 */
	public function __construct($file = '', $sheetNum = 1)
	{
		$this->set_file($file);
		$this->use_sheet($sheetNum);
	}

	/**
	 * Set filename to parse
	 *
	 * @param	string	Path to XLSX file.
	 * @return	Parser
	 */
	public function set_file($filename = null)
	{
		if (file_exists($filename)) {
			$this->file = $filename;
		} else {
			$this->errors[] = empty($filename) ? 'Missing filename' : 'Cannot find file: "'.$filename.'"';
		}

		return $this;
	}

	/**
	 * Select worksheet to parse.
	 *
	 * @param	int
	 * @return	Parser
	 */
	public function use_sheet($number = 1)
	{
		$this->sheetNum = (int) $number;

		return $this;
	}

	/**
	 * Define whether or not sheet has a header row.
	 * This is used to define the column names, which
	 * become array keys.
	 *
	 * @param	boolean
	 * @return	Parser
	 */
	public function has_header_row($value = true)
	{
		$this->has_header = (boolean) $value;

		return $this;
	}

	/**
	 * Provide an array of names to override the headers with.
	 *
	 * @param	array
	 * @return	Parser
	 */
	public function header_names($arr)
	{
		$this->header_overrides = (is_array($arr) && !empty($arr)) ? $arr : false;

		return $this;
	}

	/**
	 * Parse the spreadsheet, returning the success of the parse.
	 *
	 * @return	boolean	Successfulness of parse
	 */
	public function parse()
	{
		if ($this->file) {
			$success = $this->extract() && $this->load_files() && $this->process();
			$this->cleanup();
			return $success;
		} else {
			return false;
		}
	}

	/**
	 * Get the string path to the sharedStrings file.
	 *
	 * @return	string
	 */
	public function get_sharedStrings_file()
	{
		return $this->get_tmp_dir() . '/xl/sharedStrings.xml';
	}

	/**
	 * Get the string path to the selected worksheet's XML file.
	 *
	 * @return	string
	 */
	public function get_sheet_file()
	{
		return $this->get_tmp_dir() . '/xl/worksheets/sheet' . $this->sheetNum . '.xml';
	}

	/**
	 * Get the temporary directory where this file's contents
	 * will be extracted.
	 *
	 * @return	string
	 */
	public function get_tmp_dir()
	{
		return $this->file ? $this->base_tmp_dir . '/' . basename($this->file, '.xlsx') : '';
	}

	/**
	 * Set the directory to store temporary files.
	 * Directory must exist.
	 *
	 * @param	string	Path to temporary directory.
	 */
	public function set_tmp_dir($dir)
	{
		if (is_dir($dir))
			$this->base_tmp_dir = $dir;
	}

	/**
	 * @return	array
	 */
	public function get_parsed()
	{
		return $this->parsed;
	}

	/**
	 * Get error messages.
	 *
	 * @return	array
	 */
	public function get_errors()
	{
		return $this->errors;
	}

	/**
	 * Extract the XLSX file into a temporary directory.
	 *
	 * @return	boolean
	 */
	protected function extract()
	{
		try {
			$zip = new \ZipArchive();
			$zip->open($this->file);
			$zip->extractTo($this->get_tmp_dir());
			return true;
		} catch (\Exception $e) {
			$this->errors[] = "Failed to unzip file: " . $e->getMessage();
			return false;
		}
	}

	/**
	 * Load the important XML files for processing.
	 *
	 * @return	boolean
	 */
	protected function load_files()
	{
		$sharedStrings = $this->get_sharedStrings_file();
		$worksheet = $this->get_sheet_file();

		if (file_exists($worksheet)) {
			$this->sheet = simplexml_load_file($worksheet);
			$this->strings = simplexml_load_file($sharedStrings);
			return true;
		} else {
			$this->errors[] = 'Cannot find worksheet: ' . $this->sheetNum;
			return false;
		}
	}

	/**
	 * Process the rows into '$this->parsed'
	 *
	 * @return	boolean
	 */
	protected function process()
	{
		// Parse the rows
		$xlrows = $this->sheet->sheetData->row;

		// Process each row
		foreach ($xlrows as $xlrow)
		{
			$row = array();

			// In each row, grab its value
			foreach ($xlrow->c as $cell)
			{
				$row[row_index($cell)] = $this->process_cell($cell);
			}

			if ($this->needs_headers()) {
				$this->set_headers($row);
			} else {
				$this->parsed[] = $this->apply_headers($row);
			}
		}

		return true;
	}

	/**
	 * Process a cell.
	 *
	 * @param	\SimpleXMLElement
	 * @return	array
	 */
	protected function process_cell($cell)
	{
		$v = (string) $cell->v;

		// If it has a 't' (type) of 's' (string), use the value to look up string value.
		if (isset($cell['t']) && $cell['t'] == 's') {
			$s = array();
			$si = $this->strings->si[(int) $v];

			$si->registerXPathNamespace('n', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

			// Concatenate all of the 't' (text?) node values
			foreach($si->xpath('.//n:t') as $t)
			{
				$s[] = (string) $t;
			}

			$v = implode($s);
		}

		return $v;
	}

	/**
	 * Determine if the parser still needs headers.
	 * Only for use with sheets that have a header
	 * row.
	 *
	 * @return	boolean
	 */
	protected function needs_headers()
	{
		return $this->has_header && count($this->headers) === 0;
	}

	/**
	 * Whether or not to use default header names (e.g. a, b, c, d...)
	 *
	 * @return	boolean
	 */
	protected function is_using_default_headers()
	{
		return !$this->has_header;
	}

	/**
	 * Apply headers to a row.
	 *
	 * @param	array	Array of cells
	 * @return	array	Array of cells mapped to keys
	 */
	protected function apply_headers($row)
	{
		$composed = array();

		if ($this->is_using_default_headers()) {
			foreach($row as $idx => $row_value)
				$composed[alphacize($idx)] = $row_value;
		} else {
			foreach($this->headers as $idx => $colname)
				$composed[$colname] = isset($row[$idx]) ? $row[$idx] : '';
		}

		return $composed;
	}

	/**
	 * Set the first row to be the headers.
	 *
	 * @param	array
	 */
	protected function set_headers($arr)
	{
		$headers = array_map(function($str) {
			return preg_replace('/[\s]+/', '_', strtolower(trim($str)));
		}, $arr);

		if ($this->header_overrides)
			foreach ($this->header_overrides as $idx => $str)
				$headers[column_index($idx)] = $str;

		$this->headers = $headers;
	}

	/**
	 * Cleanup extracted files from temporary directory.
	 */
	protected function cleanup()
	{
		recursive_rm($this->get_tmp_dir());
	}
}
