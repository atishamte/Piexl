<?php

namespace atishamte\Piexl;

use PhpOffice\PhpSpreadsheet\IOFactory;

/**
 * Class Excel
 *
 * @package atishamte\Piexl
 */
class Excel
{
	/**
	 * @param array $config Configuration parameters <br>
	                       $config = [ <br>
	                           'fileNames' => 'SampleData.xlsx',           // String or Array of file names <br>
	                           'setFirstRecordAsKeys' => true,             // Set column keys as a index of each element <br>
	                           'setIndexSheetByName' => true,              // Set worksheet name as a index of sheet data array <br>
	                           'getOnlySheet'        => 'worksheetname',   // Get data of particular worksheet <br>
	                           'getOnlySheetNames'   => true,              // Get only worksheet names as a array <br>
	                           'getSheetRangeInfo'   => true,              // Get range of filled cells in worksheets <br>
	                       ]; <br>
	                       $ExcelData = Excel::import($config);
	 *
	 * @return array|bool
	 */
	public static function import($config = [])
	{
		// return false if configuration not set
		if (empty($config))
		{
			return false;
		}

		// convert single file name into array
		if (!is_array($config['fileNames']))
		{
			$config['fileNames'] = [$config['fileNames']];
		}

		// allowed excel file formats only
		$allowedFileType = ['Excel2007', 'Excel5'];
		foreach ($config['fileNames'] as $fileName)
		{
			$valid = false;
			foreach ($allowedFileType as $type)
			{
				$reader = IOFactory::createReader($type);
				if ($reader->canRead($fileName))
				{
					$valid = true;
					break;
				}
			}
			if (!$valid)
			{
				return false;
			}
		}
		$fileCount = 1;
		$returnData = [];

		// iterating over file names
		foreach ($config['fileNames'] as $fileName)
		{
			$data = [];

			// creating the reader object to read individual file
			$inputFileType = IOFactory::identify($fileName);
			$reader = IOFactory::createReader($inputFileType);

			// if getOnlySheetNames set to true then return only sheet names
			if (isset($config['getOnlySheetNames']) && $config['getOnlySheetNames'] === true)
			{
				$data = $reader->listWorksheetNames($fileName);
			}

			// if getSheetRangeInfo set to true then return only sheet data range information
			else if (isset($config['getSheetRangeInfo']) && $config['getSheetRangeInfo'] == true)
			{
				$worksheetData = $reader->listWorksheetInfo($fileName);
				foreach ($worksheetData as $worksheet)
				{
					$data[$worksheet['worksheetName']] = [
						'totalRows'    => $worksheet['totalRows'],
						'totalColumns' => $worksheet['totalColumns'],
						'cellRange'    => 'A1:' . $worksheet['lastColumnLetter'] . $worksheet['totalRows'],
					];
				}
			}
			else
			{
				// if getOnlySheet set and string passed then concert into array
				if (isset($config['getOnlySheet']))
				{
					if (!is_array($config['getOnlySheet']))
					{
						$config['getOnlySheet'] = [$config['getOnlySheet']];
					}
				}
				$spreadsheets = $reader->load($fileName);
				$sheets = $spreadsheets->getAllSheets();
				$sheetCount = 0;

				// iterate over all sheets of workbook
				foreach ($sheets as $sheet)
				{

					// if getOnlySheet set then check whether current sheet belongs to required sheets
					if (isset($config['getOnlySheet']) && !in_array(trim($sheet->getTitle()), $config['getOnlySheet']))
					{
						continue;
					}
					$sheetData = [];

					// fetch the worksheet by name to iterate over it
					$objWorksheet = $spreadsheets->setActiveSheetIndexByName($sheet->getTitle());

					// check whether column name as key parameter set or not
					$columnTaken = (isset($config['setFirstRecordAsKeys']) && $config['setFirstRecordAsKeys']) ? true : false;
					$columnNames = [];

					// iterate over each row of current sheet
					foreach ($objWorksheet->getRowIterator() as $row)
					{
						$cellIterator = $row->getCellIterator();
						$cellIterator->setIterateOnlyExistingCells(true); // This loops all cells,
						$rowData = [];
						$columnCounter = 0;

						// iterate over each cell of current row
						foreach ($cellIterator as $cell)
						{
							// if columns as a key is set to true then collect in other array else assign to data array
							if ($columnTaken)
							{
								$columnNames[] = $cell->getValue();
							}
							else
							{
								$rowData[(count($columnNames) > 0) ? $columnNames[$columnCounter] : $cell->getColumn()] = $cell->getValue();
								$columnCounter++;
							}
						}
						if (!$columnTaken)
						{
							$sheetData[] = $rowData;
						}
						$columnTaken = false;
					}

					// assign sheet name if setIndexSheetByName set to true else assign index as a key
					$data[(isset($config['setIndexSheetByName']) && $config['setIndexSheetByName']) ? $sheet->getTitle() : $sheetCount] = $sheetData;
					$sheetCount++;
				}
			}

			// append all sheet data to single file data
			$returnData['file' . $fileCount] = $data;
			$fileCount++;
		}

		// if only sing filename passed then return data without file index else return as file indexing numbered data
		return (count($config['fileNames']) > 1) ? $returnData : $returnData[key($returnData)];
	}
}