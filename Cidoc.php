<?php if ( ! defined('BASEPATH')) exit('No direct script access allowed');

require FCPATH.'vendor'.DIRECTORY_SEPARATOR.'autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

Dompdf\Autoloader::register();
use Dompdf\Dompdf;

class Cidoc {
	
	private $header;
	private $body;
	private $data;
	private $option;
	private $style;

	private $doc_ex;

	public function __construct(){
		$CI =& get_instance();
		$CI->load->helper('url');
	}

	public function set($column, $data = array(), $option = array()){

		foreach($column as $c){
			
			$this->header[]	= $c['label'];
			$this->body[]	= $c['data'];

		}

		$this->data 	= $data;
		$this->option	= $option;

		$this->processing_setup();

		return $this;
	}
	
	public function createExcel(){

		$excel = $this->processing_sheet();

		$filename = 'untitle';
		if(isset($this->option['title'])){
			$filename = url_title($this->option['title'], '_', TRUE);
		}
		
		header('Content-Type: application/vnd.ms-excel');
		header('Content-Disposition: attachment;filename="'.$filename.'.xls"');
				
		$writer = IOFactory::createWriter($excel, 'Xls');
		$writer->save('php://output');
	}

	public function createHTML($return = false){
		
		$excel = $this->processing_sheet();

		$excel->getActiveSheet()->getPageMargins()->setTop(0.2);
		$excel->getActiveSheet()->getPageMargins()->setRight(0.2);
		$excel->getActiveSheet()->getPageMargins()->setLeft(0.2);
		$excel->getActiveSheet()->getPageMargins()->setBottom(0.2);
		$excel->getActiveSheet()->getPageSetup()->setFitToWidth(1);
		$excel->getActiveSheet()->getPageSetup()->setFitToHeight(0);

		$filename = 'untitle';
		if(isset($this->option['title'])){
			$filename = url_title($this->option['title'], '_', TRUE);
		}

		$writer = new \PhpOffice\PhpSpreadsheet\Writer\Html($excel);
		$html = $writer->generateHTMLHeader();
		$html .= $this->getStyleHTML();
		$html .= $writer->generateSheetData();
		$html .= $writer->generateHTMLFooter();

		if ($return) {
			return $html;
		}

		echo $html;
	}

	public function createPDF(){

		$filename = 'untitle';
		if(isset($this->option['title'])){
			$filename = url_title($this->option['title'], '_', TRUE);
		}

		$html = $this->createHTML(true);

		$dompdf = new Dompdf();
		$dompdf->loadHtml($html);
		$dompdf->setPaper('A4', 'landscape');
		$dompdf->render();
		$dompdf->stream($filename);
	}

	public function getStyleHTML(){
		
		$style	= array(
			'table_header' => 'border: 1px solid gray; padding: 5px;',
			'table_body' => 'border: 1px solid gray; padding: 5px;',
			'summary' => 'border: 1px solid gray; padding: 5px;',
		);

		$return	= '<style>';
		$return .= 'table{ width: 100%; border-collapse: collapse; }';

		foreach($this->style as $k => $s){
			if(isset($style[$k])){
				$return .= implode(',', $s).'{'.$style[$k].'}';
			}
		}
		$return .= '</style>';

		return $return;
	}

	public function setStyleClass($type, $rownumber){
		return $this->style[$type][] = '.row'.($rownumber-1).' td';
	}

	public function processing_setup(){

		$column			= $this->header;
		$data_body		= $this->body;
		$alpha			= array();
		$alpha_first	= 'A';
		$alpha_last		= '';
		$column_pos		= array();

		for($i=1; $i<=count($column); $i++){
			$alpha[$i] = Coordinate::stringFromColumnIndex($i);
			$alpha_last = $alpha[$i];

			$column_pos[$data_body[$i-1]] = $alpha[$i];
		}

		$this->doc_ex['index_alpha']	= $alpha;
		$this->doc_ex['first_alpha']	= $alpha_first;
		$this->doc_ex['last_alpha']		= $alpha_last;
		$this->doc_ex['index_total']	= count($column);
		$this->doc_ex['column_pos']		= $column_pos;

		$data			= $this->data;
		$formatted_data	= array();
		if(!empty($data)){
			foreach($data as $d){
				$result = array_intersect_key($d, array_flip($this->body));
				$formatted_data[] = array_replace(array_flip($this->body), $result);
			}
		}

		$this->data = $formatted_data;
	}

	public function processing_sheet(){

		$prop			= $this->doc_ex;
		$row_counter	= 1;
		$data			= $this->data;
		$merged			= (isset($this->option['merged']) && is_array($this->option['merged'])) ? $this->option['merged'] : false;

		$excel	= new Spreadsheet();
		$sheet	= $excel->setActiveSheetIndex(0);

		$this->get_header_excel($sheet, $prop, $row_counter);
		$this->get_filter_info_excel($sheet, $prop, $row_counter);
		
		$row_counter++;

		// header table
		
		$cell_body = $prop['first_alpha'].$row_counter;
		foreach($prop['index_alpha'] as $k => $a){
			$sheet->setCellValue($a.$row_counter, $this->header[$k-1]);
			$this->setStyleClass('table_header', $row_counter);
		}
		
		$row_counter++;
		$mark_group = '';
		$cell_group = array();
		$cell_grouped = array();
		
		// body table
		foreach($data as $d){
			
			$a = $prop['first_alpha'];
			
			foreach($d as $dk => $dv){
				if(isset($d[$merged['by']]) && !in_array($dk, $merged['except'])){

					if($mark_group !== $d[$merged['by']]){
						$mark_group = $d[$merged['by']];
						
						if(array_key_exists($a.'_'.$mark_group, $cell_group)){
							$mark_group .= date('Ymdhis');
						}
						
						$cell_group[$a.'_'.$mark_group][] = $a.$row_counter;
					}else{
						$cell_group[$a.'_'.$mark_group][] = $a.$row_counter;
					}
				}
				
				$this->setStyleClass('table_body', $row_counter);
				$sheet->setCellValue($a.$row_counter, $dv);
				$a++;
			}
			
			$row_counter++;
		}

		if(!empty($cell_group)){
			foreach($cell_group as $g){

				$g_first = reset($g);
				$g_last = end($g);

				if($g_first && $g_last && $g_first !== $g_last){
					$sheet->mergeCells($g_first.':'.$g_last);
					$sheet->getStyle($g_first.':'.$g_last)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
				}
			}
		}

		$this->get_summary_excel($sheet, $prop, $row_counter);


		$cell_body .= ':'.$prop['last_alpha'].($row_counter-1);

		$styleArray = array(
            'borders' => array(
                'allBorders' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ),
            ),
        );

        $sheet->getStyle($cell_body)->applyFromArray($styleArray);

		foreach($prop['index_alpha'] as $a){
      $sheet->getColumnDimension($a)->setAutoSize(true);
		}
		
		return $excel;
	}

	public function get_header_excel(&$sheet, $prop, &$row_counter){
		if(isset($this->option['title']) && !empty($this->option['title'])){
			$sheet->setCellValue($prop['first_alpha'].$row_counter, $this->option['title']);
			$sheet->mergeCells($prop['first_alpha'].$row_counter.':'.$prop['last_alpha'].$row_counter);
			$this->setStyleClass('table_title', $row_counter);
			$row_counter++;
		}
	}

	public function get_filter_info_excel(&$sheet, $prop, &$row_counter){

		if(!isset($this->option['filter']) || !is_array($this->option['filter'])){
			return;
		}

		$row_counter++;
		
		foreach($this->option['filter'] as $k => $f){
			$this->setStyleClass('filter', $row_counter);
			$sheet->setCellValue($prop['first_alpha'].$row_counter++, $f['label'].' : '.$f['value']);
		}

		$row_counter++;

	}

	public function get_summary_excel(&$sheet, $prop, &$row_counter)
	{
		if(!isset($this->option['summary']) || !is_array($this->option['summary'])){
			return;
		}

		foreach($this->option['summary'] as $f){
			$this->setStyleClass('summary', $row_counter);
			$sheet->setCellValue($prop['first_alpha'].$row_counter, $f['label']);

			if(!is_array($f['column'])){
				if(isset($prop['column_pos'][$f['column']])){
					$sheet->setCellValue($prop['column_pos'][$f['column']].$row_counter, $f['value']);
				}
			}else{
				foreach($f['column'] as $ck => $cv){
					$sheet->setCellValue($prop['column_pos'][$ck].$row_counter, $cv);
				}
			}
			
			$row_counter++;
		}
	}
}