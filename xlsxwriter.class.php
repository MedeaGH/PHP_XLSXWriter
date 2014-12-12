<?php
/*
 * @license MIT License
 * */

if (!class_exists("ZipArchive")) { Throw New Exception("ZipArchive not found"); }

class XLSXWriter
{
	//------------------------------------------------------------------
	//http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
	const EXCEL_2007_MAX_ROW = 1048576; 
	const EXCEL_2007_MAX_COL = 16384;

	//------------------------------------------------------------------
	protected $author              = "Author";
	protected $sheets              = array();
	protected $sharedStrings       = array(); // unique set
	protected $sharedStringCount   = 0;       // count of non-unique references to the unique set
	protected $tempFiles           = array();

	# Default values
	protected $defaultFontname     = "Calibri";
	protected $defaultFontsize     = 11;

	protected $defaultColWidth     = 20;

	# Encoding
	protected $encodeToUTF8        = false;

	# Timezone
	protected $timeZone            = "Europe/Paris";

	protected $styles              = array(); // style  for each sheet / row / cols
	protected $formats             = array(); // format for each sheet / cols

	protected $currentSheet        = "";

	/**
	* Styles
	*/
	protected $_numFmts = array();
	protected $_fonts   = array();
	protected $_fills   = array();
	protected $_borders = array();

	protected $_styles  = array();
	
	public function __construct()
	{
		if (!ini_get("date.timezone"))
		{
			date_default_timezone_set($this->timeZone);
		}

		/**
		* Fonts init.
		*/
		$this->_fonts[] = array (
			"name" => $this->defaultFontname,
			"size" => $this->defaultFontsize,
			"opts" => array(),
		);

		$this->_fonts[] = array (
			"name" => $this->defaultFontname,
			"size" => $this->defaultFontsize,
			"opts" => array( "bold" ),
		);

		$this->_fonts[] = array (
			"name" => $this->defaultFontname,
			"size" => $this->defaultFontsize,
			"opts" => array( "italic" ),
		);

		$this->_fonts[] = array (
			"name" => $this->defaultFontname,
			"size" => $this->defaultFontsize,
			"opts" => array( "bold", "italic"),
		);

		foreach ($this->_fonts as $i => $font)
		{
			$this->_fonts[$i]["md5"] = md5(serialize($font));
		}

		/**
		* Fills init.
		*/
		$this->_fills[] = array (
			"type" => "none",
		);

		$this->_fills[] = array (
			"type" => "grey125",
		);

		$this->_fills[] = array (
			"type"    => "solid",
			"fgColor" => "00F2F2F2",
		);

		$this->_fills[] = array (
			"type"    => "solid",
			"fgColor" => "00FF8C8C",
		);

		foreach ($this->_fills as $i => $fill)
		{
			$this->_fills[$i]["md5"] = md5(serialize($fill));
		}

		/**
		* Borders init.
		*/
		$this->_borders[] = array (
			"left"     => array(),
			"right"    => array(),
			"top"      => array(),
			"bottom"   => array(),
			"diagonal" => array(),
		);

		$this->_borders[] = array (
			"left"      => array (
				"style" => "thin",
				"color" => 64
			),
			"right"     => array (
				"style" => "thin",
				"color" => 64
			),
			"top"       => array (
				"style" => "thin",
				"color" => 64
			),
			"bottom"    => array (
				"style" => "thin",
				"color" => 64
			),
			"diagonal"  => array (
				"style" => "thin",
				"color" => 64
			),
		);

		foreach ($this->_borders as $i => $border)
		{
			$this->_borders[$i]["md5"] = md5(serialize($border));
		}

		/**
		* Styles init.
		*/
		$this->getStyleID();
	}

	public function __destruct()
	{
		if (!empty($this->tempFiles))
		{
			foreach($this->tempFiles as $tempFile)
			{
				@unlink($tempFile);
			}
		}
	}

	public function addFont($name, $size, $opts = array())
	{
		$font = array (
			"name" => $name,
			"size" => $size,
			"opts" => $opts,
		);

		$font["md5"] = md5(serialize($font));

		$this->_fonts[] = $font;

		return (count($this->_fonts) - 1);
	}

	public function setTimeZone($timeZone)
	{
		$this->timeZone = $timeZone;
		date_default_timezone_set($this->timeZone);
	}

	public function setAuthor($author = "")
	{
		$this->author = $author;
	}

	public function setUTF8($value)
	{
		$this->encodeToUTF8 = (is_bool($value) ? $value : true);
	}

	protected function getStyleID($styleData = array(), $format = "string")
	{
		if (isset($styleData["font"])   === false) $styleData["font"]   = 0;
		if (isset($styleData["fill"])   === false) $styleData["fill"]   = 0;
		if (isset($styleData["border"]) === false) $styleData["border"] = 0;
		if (isset($styleData["numFmt"]) === false) $styleData["numFmt"] = 0;
		if (isset($styleData["halign"]) === false) $styleData["halign"] = "";
		if (isset($styleData["valign"]) === false) $styleData["valign"] = "";

		$styleData["format"] = $format;

		$key = md5(serialize($styleData));

		$currentID = count($this->_styles);

		if (isset($this->_styles[$key]) === false)
		{
			$style = array (
				"id"  => $currentID,
				"xml" => ""
			);

			$applyFont      = (!empty($styleData["font"]) ? true : false);
			$applyFill      = (!empty($styleData["fill"]) ? true : false);
			$applyBorder    = (!empty($styleData["border"]) ? true : false);
			$applyAlignment = (!empty($styleData["halign"]) || !empty($styleData["valign"]) ? true : false);

			$xml  = '<xf numFmtId="' . $styleData["numFmt"]. '" fontId="' . $styleData["font"] . '" fillId="' . $styleData["fill"] . '" borderId="' . $styleData["border"] . '" xfId="0"';
			
			if ($applyFont === true)
				$xml .= ' applyFont="1"';

			if ($applyFill === true)
				$xml .= ' applyFill="1"';

			if ($applyBorder === true)
				$xml .= ' applyBorder="1"';

			if ($applyAlignment === true)
				$xml .= ' applyAlignment="1"';

			$xml .= '>';

			if ($applyAlignment === true)
			{
				$xml .= "<alignment";

				if (!empty($styleData["halign"]))
					$xml .= ' horizontal="' . $styleData["halign"] . '"';

				if (!empty($styleData["valign"]))
					$xml .= ' vertical="' . $styleData["valign"] . '"';
			
				$xml .= "/>";
			}

			$xml .= '</xf>';

			$style["xml"] = $xml;

			$this->_styles[$key] = $style;

			return $currentID;
		}
		else
		{
			return $this->_styles[$key]["id"];
		}
	}

	/**
	* $type 
	*/
	public function setStyles($sheetName, $type, $style)
	{
		$this->styles[$sheetName][$type] = $style;
	}

	public function setColStyle($sheetName, $col, $style)
	{
		$this->styles[$sheetName]["cols"][$col] = $style;
	}

	public function setColFormat($sheetName, $col, $format)
	{
		$this->formats[$sheetName][$col] = $format;
	}

	protected function tempFilename()
	{
		$filename = tempnam(sys_get_temp_dir(), "xlsx_writer_");
		$this->tempFiles[] = $filename;

		return $filename;
	}

	public function writeToStdOut()
	{
		$tempFile = $this->tempFilename();
		self::writeToFile($tempFile);
		readfile($tempFile);
	}

	public function writeToString()
	{
		$tempFile = $this->tempFilename();
		self::writeToFile($tempFile);
		$string = file_get_contents($tempFile);
		return $string;
	}

	public function writeToFile($filename)
	{
		foreach($this->sheets as $sheetName => $sheet)
		{
			self::finalizeSheet($sheetName); //making sure all footers have been written
		}

		@unlink($filename); //if the zip already exists, overwrite it

		$zip = new ZipArchive();
		if (empty($this->sheets))                       { self::log("Error in ".__CLASS__."::".__FUNCTION__.", no worksheets defined."); return; }
		if (!$zip->open($filename, ZipArchive::CREATE)) { self::log("Error in ".__CLASS__."::".__FUNCTION__.", unable to create zip."); return; }
		
		$zip->addEmptyDir("docProps/");
		$zip->addFromString("docProps/app.xml" , self::buildAppXML() );
		$zip->addFromString("docProps/core.xml", self::buildCoreXML());

		$zip->addEmptyDir("_rels/");
		$zip->addFromString("_rels/.rels", self::buildRelationshipsXML());

		$zip->addEmptyDir("xl/worksheets/");
		foreach($this->sheets as $sheet)
		{
			$zip->addFile($sheet->filename, "xl/worksheets/" . $sheet->xmlName );
		}

		if (!empty($this->sharedStrings))
		{
			$zip->addFile($this->writeSharedStringsXML(), "xl/sharedStrings.xml" );
		}
		$zip->addFromString("xl/workbook.xml", self::buildWorkbookXML() );
		$zip->addFile($this->writeStylesXML(), "xl/styles.xml" );

		$zip->addFromString("[Content_Types].xml", self::buildContentTypesXML() );

		$zip->addEmptyDir("xl/_rels/");
		$zip->addFromString("xl/_rels/workbook.xml.rels", self::buildWorkbookRelsXML() );

		$zip->close();
	}

	public function initializeSheet($sheetName, $sheetOptions = array())
	{
		//if already initialized
		if ($this->currentSheet == $sheetName || isset($this->sheets[$sheetName]))
			return;

		$sheetFilename = $this->tempFilename();
		$sheetXMLName = "sheet" . (count($this->sheets) + 1) . ".xml";

		$this->sheets[$sheetName] = (object)array (
			"filename"           => $sheetFilename,
			"sheetName"          => $sheetName,
			"xmlName"            => $sheetXMLName,
			"rowCount"           => 0,
			"fileWriter"         => new XLSXWriter_BuffererWriter($sheetFilename),
			"colsWidth"          => array(),
			"maxCellTagStart"    => 0,
			"maxCellTagEnd"      => 0,
			"finalized"          => false,
		);

		$colSize = (isset($sheetOptions["colSize"]) ? $sheetOptions["colSize"] : $this->defaultColWidth);

		$sheet = &$this->sheets[$sheetName];
		$tabSelected = count($this->sheets) == 1 ? "true" : "false";

		$sheet->fileWriter->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
		$sheet->fileWriter->write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');

		$sheet->fileWriter->write('<dimension ref="A1"/>');

		$sheet->fileWriter->write(  '<sheetViews>');
		$sheet->fileWriter->write(    '<sheetView workbookViewId="0"/>');
		$sheet->fileWriter->write(  '</sheetViews>');
		$sheet->fileWriter->write(  '<cols>');
		$sheet->fileWriter->write(    '<col min="1" max="' . self::EXCEL_2007_MAX_COL . '" style="0" customWidth="1" width="' . $colSize . '"/>');
		$sheet->fileWriter->write(  '</cols>');
		$sheet->fileWriter->write(  '<sheetData>');
	}

	public function writeSheetHeader($sheetName, array $row, array $style = array())
	{
		if (empty($sheetName) || empty($row))
			return;

		self::initializeSheet($sheetName);
		$sheet = &$this->sheets[$sheetName];

		$rowStyle = (isset($this->styles[$sheetName]["header"]) ? $this->styles[$sheetName]["header"] : $style);

		$sheet->fileWriter->write('<row r="' . ($sheet->rowCount + 1) . '">');

		foreach ($row as $columnNumber => $value)
		{
			$cellOptions = array (
				"style"  => $rowStyle,
				"format" => "string",
			);

			$this->writeCell($sheet->fileWriter, $sheet->rowCount, $columnNumber, $value, $cellOptions);
		}

		$sheet->fileWriter->write('</row>');
		$sheet->rowCount++;

		$this->currentSheet = $sheetName;
	}

	public function writeSheetRow($sheetName, array $row, array $style = array())
	{
		if (empty($sheetName) || empty($row))
			return;

		self::initializeSheet($sheetName);
		$sheet = &$this->sheets[$sheetName];

		$rowStyle = (isset($this->styles[$sheetName]["row"])  ? $this->styles[$sheetName]["row"] : $style);

		$sheet->fileWriter->write('<row r="' . ($sheet->rowCount + 1) . '">');

		foreach ($row as $columnNumber => $value)
		{
			$cellOptions = array (
				"style"  => $rowStyle,
				"format" => "string",
			);

			if (isset($this->styles[$sheetName]["cols"][$columnNumber]))
			{
				$cellOptions["style"] = $this->styles[$sheetName]["cols"][$columnNumber];
			}

			if (isset($this->formats[$sheetName][$columnNumber]))
			{
				$cellOptions["format"] = $this->formats[$sheetName][$columnNumber];
			}

			$this->writeCell($sheet->fileWriter, $sheet->rowCount, $columnNumber, $value, $cellOptions);
		}

		$sheet->fileWriter->write('</row>');
		$sheet->rowCount++;

		$this->currentSheet = $sheetName;
	}
	
	protected function finalizeSheet($sheetName)
	{
		if (empty($sheetName) || $this->sheets[$sheetName]->finalized)
			return;

		$sheet = &$this->sheets[$sheetName];

		$sheet->fileWriter->write(    '</sheetData>');
		$sheet->fileWriter->write('</worksheet>');

		$sheet->fileWriter->close();

		$sheet->finalized = true;
	}

	public function writeSheet(array $data, $sheetName = "")
	{
		$sheetName = (empty($sheetName) ? "Sheet1" : $sheetName);
		$data      = (empty($data) ? array(array("")) : $data);

		foreach ($data as $i => $row)
		{
			$this->writeSheetRow($sheetName, $row);
		}

		$this->finalizeSheet($sheetName);
	}

	protected function writeCell(XLSXWriter_BuffererWriter &$file, $rowNumber, $columnNumber, $value, $cellOptions = array())
	{
		$cell = self::xlsCell($rowNumber, $columnNumber);
		
		$styleID = $this->getStyleID($cellOptions["style"], $cellOptions["format"]);

		if (!is_scalar($value) || $value == "") // objects, array, empty
		{
			$file->write('<c r="' . $cell . '"' . ($styleID ? ' s="' . $styleID . '"' : '') . '/>');
		}
		elseif ($cellOptions["format"] == "date")
		{
			$file->write('<c r="' . $cell . '"' . ($styleID ? ' s="' . $styleID . '"' : '') . ' t="n"><v>' . intval(self::convertDateTime($value)) . '</v></c>');
		}
		elseif ($cellOptions["format"] == "datetime")
		{
			$file->write('<c r="' . $cell . '"' . ($styleID ? ' s="' . $styleID . '"' : '') . ' t="n"><v>' . self::convertDateTime($value) . '</v></c>');
		}
		elseif (is_string($value) === false)
		{
			$file->write('<c r="' . $cell . '"' . ($styleID ? ' s="' . $styleID . '"' : '') . ' t="n"><v>' . ($value * 1 ) . '</v></c>');//int,float, etc
		}
		elseif ($value{0} != "0" && filter_var($value, FILTER_VALIDATE_INT)) // excel wants to trim leading zeros
		{ 
			$file->write('<c r="' . $cell . '"' . ($styleID ? ' s="' . $styleID . '"' : '') . ' t="n"><v>' . ($value) . '</v></c>');//numeric string
		} 
		elseif ($value{0} == "=")
		{
			$file->write('<c r="' . $cell . '"' . ($styleID ? ' s="' . $styleID . '"' : '') . ' t="s"><f>' . self::xmlSpecialChars($value) . '</f></c>');
		} 
		elseif ($value !== "")
		{
			$value = $this->utf8($value);

			$file->write('<c r="' . $cell . '"' . ($styleID ? ' s="' . $styleID . '"' : '') . ' t="s"><v>' . self::xmlSpecialChars($this->setSharedString($value)) . '</v></c>');
		}
	}

	protected function writeStylesXML()
	{
		$temporaryFilename = $this->tempFilename();

		$file = new XLSXWriter_BuffererWriter($temporaryFilename);

		$file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
		$file->write('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');

//		$file->write('<numFmts count="4">');
//		$file->write(	'<numFmt formatCode="GENERAL" numFmtId="164"/>');
//		$file->write(	'<numFmt formatCode="[$$-1009]#,##0.00;[RED]\-[$$-1009]#,##0.00" numFmtId="165"/>');
//		$file->write(	'<numFmt formatCode="YYYY-MM-DD\ HH:MM:SS" numFmtId="166"/>');
//		$file->write(	'<numFmt formatCode="YYYY-MM-DD" numFmtId="167"/>');
//		$file->write('</numFmts>');

		/**
		* Fonts
		*/
		$xml = $this->getStylesFontsXML();
		$file->write($xml);

		/**
		* Fills
		*/
		$xml = $this->getStylesFillsXML();
		$file->write($xml);

		/**
		* Borders
		*/
		$xml = $this->getStylesBordersXML();
		$file->write($xml);

		/**
		* base
		*/
		$file->write('<cellStyleXfs count="1">');
		$file->write(	'<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>');
		$file->write('</cellStyleXfs>');

		/**
		* Styles : the true one
		*/
		$xml = '<cellXfs count="' . count($this->_styles) . '">';

		foreach ($this->_styles as $style)
		{
			$xml .= $style["xml"];
		}

		$xml .= '</cellXfs>';

		$file->write($xml);

		/**
		* base
		*/
		$file->write('<cellStyles count="1">');
		$file->write(	'<cellStyle name="Normal" xfId="0" builtinId="0"/>');
		$file->write('</cellStyles>');

		$file->write('</styleSheet>');
		$file->close();

		return $temporaryFilename;
	}

	protected function setSharedString($v)
	{
		if (isset($this->sharedStrings[$v]))
		{
			$stringValue = $this->sharedStrings[$v];
		}
		else
		{
			$stringValue = count($this->sharedStrings);
			$this->sharedStrings[$v] = $this->utf8($stringValue);
		}
		
		$this->sharedStringCount++;

		return $stringValue;
	}

	protected function getStylesFontsXML()
	{
		$xml  = '';
		$xml .= '<fonts count="' . count($this->_fonts) . '">';

		foreach ($this->_fonts as $font)
		{
			$xml .= '<font>';
			$xml .= '<name val="' . $font["name"] . '"/>';

			if (in_array("bold", $font["opts"]) === true)
			{
				$xml .= '<b/>';
			}

			if (in_array("italic", $font["opts"]) === true)
			{
				$xml .= '<i/>';
			}

			if (in_array("underline", $font["opts"]) === true)
			{
				$xml .= '<u/>';
			}

			if (in_array("double_underline", $font["opts"]) === true)
			{
				$xml .= '<u val="double"/>';
			}

			$xml .= '<family val="2"/>';
			$xml .= '<sz val="' . $font["size"] . '"/>';

			$xml .= '</font>';
		}

		$xml .= '</fonts>';

		return $xml;
	}

	protected function getStylesFillsXML()
	{
		$xml  = '';
		$xml .= '<fills count="' . count($this->_fills) . '">';

		foreach ($this->_fills as $fill)
		{
			$xml .= '<fill>';

			$xml .= '<patternFill patternType="' . $fill["type"] . '">';

			if (!empty($fill["fgColor"]))
			{
				if (is_numeric($fill["fgColor"]))
				{
					$xml .= '<fgColor indexed="' . $fill["fgColor"] . '"/>';
				}
				else
				{
					$xml .= '<fgColor rgb="' . $fill["fgColor"] . '"/>';
				}
			}

			if (!empty($fill["bgColor"]))
			{
				if (is_numeric($fill["bgColor"]))
				{
					$xml .= '<bgColor indexed="' . $fill["bgColor"] . '"/>';
				}
				else
				{
					$xml .= '<bgColor rgb="' . $fill["bgColor"] . '"/>';
				}
			}

			$xml .= '</patternFill>';

			$xml .= '</fill>';
		}

		$xml .= '</fills>';

		return $xml;
	}

	protected function getStylesBordersXML()
	{
		$xml  = '';
		$xml .= '<borders count="' . count($this->_borders) . '">';

		$borderTypes = array ("left", "right", "top", "bottom", "diagonal");

		foreach ($this->_borders as $border)
		{
			$xml .= '<border>';

			foreach ($borderTypes as $borderType)
			{
				if (empty($border[$borderType]))
				{
					$xml .= '<' . $borderType . '/>';
				}
				else
				{
					$xml .= '<' . $borderType . ' style="' . $border[$borderType]["style"] . '">';

					if (!empty($border[$borderType]["color"]))
					{
						if (is_numeric($border[$borderType]["color"]))
						{
							$xml .= '<color indexed="' .$border[$borderType]["color"] . '"/>';
						}
						else
						{
							$xml .= '<color rgb="' . $border[$borderType]["color"] . '"/>';							
						}
					}

					$xml .= '</' . $borderType . '>';
				}
			}

			$xml .= '</border>';			
		}

		$xml .= '</borders>';

		return $xml;
	}

	protected function writeSharedStringsXML()
	{
		$temporaryFilename = $this->tempFilename();

		$file = new XLSXWriter_BuffererWriter($temporaryFilename, $fopenFlags = "w", $checkUTF8 = true);
		
		$file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
		$file->write('<sst count="' . ($this->sharedStringCount) . '" uniqueCount="' . count($this->sharedStrings) . '" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
		
		foreach($this->sharedStrings as $s=>$c)
		{
			$file->write('<si><t>' . self::xmlSpecialChars($s) . '</t></si>');
		}
		
		$file->write('</sst>');
		$file->close();
		
		return $temporaryFilename;
	}

	protected function buildAppXML()
	{
		$xml  = '';
		$xml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
		$xml .= '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>0</TotalTime></Properties>';

		return $xml;
	}

	protected function buildCoreXML()
	{
		$author = $this->utf8($this->author);

		$xml  = '';
		$xml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
		$xml .= '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
		$xml .= '<dcterms:created xsi:type="dcterms:W3CDTF">' . date("Y-m-d\TH:i:s.00\Z") . '</dcterms:created>';
		$xml .= '<dc:creator>' . self::xmlSpecialChars($author) . '</dc:creator>';
		$xml .= '<cp:revision>0</cp:revision>';
		$xml .= '</cp:coreProperties>';

		return $xml;
	}

	protected function buildRelationshipsXML()
	{
		$xml  = '';
		$xml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
		$xml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
		$xml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>';
		$xml .= '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>';
		$xml .= '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>';
		$xml .= "\n";
		$xml .= '</Relationships>';

		return $xml;
	}

	protected function buildWorkbookXML()
	{
		$i = 0;

		$xml  = '';
		$xml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
		$xml .= '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
		$xml .= '<fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>';
		$xml .= '<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>';
		$xml .= '<sheets>';

		foreach ($this->sheets as $sheetName => $sheet)
		{
			$sheetName = $this->utf8($sheetName);

			$xml .= '<sheet name="' . self::xmlSpecialChars($sheetName) . '" sheetId="' . ($i + 1) . '" state="visible" r:id="rId' . ($i + 2) . '"/>';
			$i++;
		}

		$xml .= '</sheets>';
		$xml .= '<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/>';
		$xml .= '</workbook>';

		return $xml;
	}

	protected function utf8($string)
	{
		return ($this->encodeToUTF8 === true ? utf8_encode($string) : $string);
	}

	protected function buildWorkbookRelsXML()
	{
		$i = 0;

		$xml  = '';
		$xml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
		$xml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
		$xml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';

		foreach($this->sheets as $sheetName => $sheet)
		{
			$xml .= '<Relationship Id="rId' . ($i + 2) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/' . ($sheet->xmlName) . '"/>';
			$i++;
		}

		if (!empty($this->sharedStrings))
		{
			$xml .= '<Relationship Id="rId' . (count($this->sheets) + 2) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>';
		}

		$xml .= "\n";
		$xml .= '</Relationships>';

		return $xml;
	}

	protected function buildContentTypesXML()
	{
		$xml  = '';
		$xml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
		$xml .= '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
		$xml .= '<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
		$xml .= '<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
		
		foreach($this->sheets as $sheetName => $sheet)
		{
			$xml .= '<Override PartName="/xl/worksheets/' . ($sheet->xmlName) . '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
		}

		if (!empty($this->sharedStrings))
		{
			$xml .= '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
		}

		$xml .= '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
		$xml .= '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
		$xml .= '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
		$xml .= '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
		$xml .= "\n";
		$xml .= '</Types>';

		return $xml;
	}

	//------------------------------------------------------------------
	/*
	 * @param $row_number int, zero based
	 * @param $column_number int, zero based
	 * @return Cell label/coordinates, ex: A1, C3, AA42
	 * */
	public static function xlsCell($rowNumber, $columnNumber)
	{
		$n = $columnNumber;

		for ($r = ""; $n >= 0; $n = intval($n / 26) - 1)
		{
			$r = chr($n % 26 + 0x41) . $r;
		}

		return $r . ($rowNumber + 1);
	}

	//------------------------------------------------------------------
	public static function log($string)
	{
		file_put_contents("php://stderr", date("Y-m-d H:i:s:").rtrim(is_array($string) ? json_encode($string) : $string)."\n");
	}
	//------------------------------------------------------------------
	public static function sanitizeFilename($filename) //http://msdn.microsoft.com/en-us/library/aa365247%28VS.85%29.aspx
	{
		$nonPrinting  = array_map("chr", range(0, 31));
		$invalidChars = array("<", ">", "?", "\"", ":", "|", "\\", "/", "*", "&");
		$allInvalids  = array_merge($nonPrinting, $invalidChars);

		return str_replace($allInvalids, "", $filename);
	}
	//------------------------------------------------------------------
	public static function xmlSpecialChars($val)
	{
		return str_replace("'", "&#39;", htmlspecialchars($val));
	}
	//------------------------------------------------------------------
	public static function arrayFirstKey(array $arr)
	{
		reset($arr);
		$firstKey = key($arr);
		return $firstKey;
	}
	//------------------------------------------------------------------
	public static function convertDateTime($dateInput) //thanks to Excel::Writer::XLSX::Worksheet.pm (perl)
	{
		$days    = 0;    # Number of days since epoch
		$seconds = 0;    # Time expressed as fraction of 24h hours in seconds
		$year = $month = $day = 0;
		$hour = $min   = $sec = 0;

		$dateTime = $dateInput;

		if (preg_match("/(\d{4})\-(\d{2})\-(\d{2})/", $dateTime, $matches))
		{
			list($junk, $year, $month, $day) = $matches;
		}
		if (preg_match("/(\d{2}):(\d{2}):(\d{2})/", $dateTime, $matches))
		{
			list($junk, $hour, $min, $sec) = $matches;

			$seconds = ( $hour * 60 * 60 + $min * 60 + $sec ) / ( 24 * 60 * 60 );
		}

		//using 1900 as epoch, not 1904, ignoring 1904 special case
		
		# Special cases for Excel.
		if ($year . "-" . $month . "-" . $day == "1899-12-31")  return $seconds      ;    # Excel 1900 epoch
		if ($year . "-" . $month . "-" . $day == "1900-01-00")  return $seconds      ;    # Excel 1900 epoch
		if ($year . "-" . $month . "-" . $day == "1900-02-29")  return 60 + $seconds ;    # Excel false leapday

		# We calculate the date by calculating the number of days since the epoch
		# and adjust for the number of leap days. We calculate the number of leap
		# days by normalising the year in relation to the epoch. Thus the year 2000
		# becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
		$epoch  = 1900;
		$offset = 0;
		$norm   = 300;
		$range  = $year - $epoch;

		# Set month days and check for leap year.
		$leap = (($year % 400 == 0) || (($year % 4 == 0) && ($year % 100)) ) ? 1 : 0;
		$mdays = array( 31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 );

		# Some boundary checks
		if ($year < $epoch || $year > 9999) return 0;
		if ($month < 1     || $month > 12)  return 0;
		if ($day < 1       || $day > $mdays[ $month - 1 ]) return 0;

		# Accumulate the number of days since the epoch.
		$days = $day;                                                 # Add days for current month
		$days += array_sum( array_slice($mdays, 0, $month - 1 ) );    # Add days for past months
		$days += $range * 365;                                        # Add days for past years
		$days += intval( ( $range ) / 4 );                            # Add leapdays
		$days -= intval( ( $range + $offset ) / 100 );                # Subtract 100 year leapdays
		$days += intval( ( $range + $offset + $norm ) / 400 );        # Add 400 year leapdays
		$days -= $leap;                                               # Already counted above

		# Adjust for Excel erroneously treating 1900 as a leap year.
		if ($days > 59) { $days++; }

		return $days + $seconds;
	}

} /* END CLASS : XLSXWriter */

class XLSXWriter_BuffererWriter
{
	protected $fd         = null;
	protected $buffer     = "";
	protected $checkUTF8  = false;

	public function __construct($filename, $fopenFlags = "w", $checkUTF8 = false)
	{
		$this->checkUTF8 = $checkUTF8;
		$this->fd = fopen($filename, $fopenFlags);
		if ($this->fd === false)
		{
			XLSXWriter::log("Unable to open $filename for writing.");
		}
	}

	public function write($string)
	{
		$this->buffer .= $string;
		if (isset($this->buffer[8191]))
		{
			$this->purge();
		}
	}

	protected function purge()
	{
		if ($this->fd)
		{
			if ($this->checkUTF8 && !self::isValidUTF8($this->buffer))
			{
				XLSXWriter::log("Error, invalid UTF8 encoding detected.");
				$this->checkUTF8 = false;
			}
			fwrite($this->fd, $this->buffer);
			$this->buffer = "";
		}
	}

	public function close()
	{
		$this->purge();
		if ($this->fd)
		{
			fclose($this->fd);
			$this->fd = null;
		}
	}

	public function __destruct() 
	{
		$this->close();
	}
	
	public function ftell()
	{
		if ($this->fd)
		{
			$this->purge();
			return ftell($this->fd);
		}
		return -1;
	}

	public function fseek($pos)
	{
		if ($this->fd)
		{
			$this->purge();
			return fseek($this->fd, $pos);
		}
		return -1;
	}

	protected static function isValidUTF8($string)
	{
		if (function_exists("mb_check_encoding"))
		{
			return mb_check_encoding($string, "UTF-8") ? true : false;
		}
		return preg_match("//u", $string) ? true : false;
	}
}

/* END CLASS : XLSXWriter_BuffererWriter */

?>