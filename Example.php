<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Example extends CI_Controller {

	public function index()
	{
		// Loading Lib
		$this->load->library('cidoc');

		// Preparing Data
		$data = [];
		for($i=0; $i<50; $i++){
			$name = 'Lorem Ipsum';
			if ($i > 11 && $i < 15) {
				$name = 'Dolor Sit';
			}

			$data[] = array(
				'id' => $i+1,
				'date' => date('d M Y'),
				'name' => $name,
				'hit' => rand(0, 100),
				'status' => 'Complete',
			);
		}

		// Preparing Table Header and Mapping Data
		$header = array(
				array('label' => 'ID.', 'data' => 'id'),
				array('label' => 'Date', 'data' => 'date'),
				array('label' => 'Name', 'data' => 'name'),
				array('label' => 'Hit', 'data' => 'hit'),
				array('label' => 'Status', 'data' => 'status'),
		);

		// Preparing Option Lib
		$option = array(
				'title' => 'User Report', // Set the title of table
				'summary' => array( // Set summary or footer table
						array(
								'label' => 'Total Hit', 
								'column' => array(
										'hit' => array_sum(array_column($data, 'hit')),
								)
						)
				),
				'filter' => array( // Set more information on header
					array('label' => 'Date', 'value' => date('d M Y')),
					array('label' => 'Status', 'value' => 'Complete'),
				),
				'merged' => array('by' => 'name', 'except' => array('id', 'hit', 'date', 'status')), // Merge unique name data vertically
		);

		$this->cidoc->set($header, $data, $option)->createHTML(); // Output HTML
		// $this->cidoc->set($header, $data, $option)->createExcel(); // Output Excel
		// $this->cidoc->set($header, $data, $option)->createPDF(); // Output PDF
	}
}
