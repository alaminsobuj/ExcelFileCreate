	 	$fileName =$fileName = 'Trialbalance_report.xlsx';;
	 	$this->load->library('excel');
	 	$objPHPExcel = new PHPExcel();
	 	$objPHPExcel->setActiveSheetIndex(0);
	 	$objPHPExcel->getActiveSheet()->getStyle('1:1')->getFont()->setBold(true);
		$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A1:L1');
		//===================== set Header=======================================
	 	$objPHPExcel->getActiveSheet()->SetCellValue('A1','Company Name Here');
	 	$objPHPExcel->getActiveSheet()->getStyle('2:2')->getFont()->setBold(true);
	 	$objPHPExcel->getActiveSheet()->SetCellValue('A2','Po Request');
		$objPHPExcel->setActiveSheetIndex(0)->mergeCells('A2:G2');

		$objPHPExcel->getActiveSheet()->SetCellValue('A3',  'Product Type');
	 	$objPHPExcel->getActiveSheet()->SetCellValue('B3',  'Item Information');
	 	$objPHPExcel->getActiveSheet()->SetCellValue('C3',  'Quantity');
		$objPHPExcel->getActiveSheet()->SetCellValue('D3',  'Rate');
	 	$objPHPExcel->getActiveSheet()->SetCellValue('E3',  'Vat Type');
	 	$objPHPExcel->getActiveSheet()->SetCellValue('F3',  'Total');

	 	$rowCount = 5;

                #==============this is row=============================================== 

	 	$objPHPExcel->getActiveSheet()->SetCellValue('A'. $rowCount,'RowFood');
	 	$objPHPExcel->getActiveSheet()->SetCellValue('B'. $rowCount,'Onion');
	 	$objPHPExcel->getActiveSheet()->SetCellValue('C'. $rowCount, 10);
	 	$objPHPExcel->getActiveSheet()->SetCellValue('D'. $rowCount, 500);
	 	$objPHPExcel->getActiveSheet()->SetCellValue('E'. $rowCount, '%');
	 	$objPHPExcel->getActiveSheet()->SetCellValue('F'. $rowCount, 510);

	// 	// dd($objPHPExcel);
		// $rowCount++;

                #==============this is row=============================================== 


	 	$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);

		$objWriter->save($fileName);
                //$objWriter->save('./application/modules/purchase/assets/excel/'.$fileName);

	 	// download file
		header("Content-Type: application/vnd.ms-excel");
		redirect(site_url().$fileName);



                ##email send 

                
                $this->email->initialize($config);
		$this->email->set_newline("\r\n");
		$this->email->set_mailtype("html");
		// $htmlContent = ReservationEmail($insert_id, $mobile);
		$this->email->from($send_email->sender);
		$this->email->to($supplier->supEmail);
		$this->email->subject('test');
		$this->email->message(
		  '<h1>Po Order Request </h1> <br>'.
		 'Company Name :'.$companyinfo->storename.'<br>'.'Phone Number:'.$companyinfo->phone.'<br>'.'Address :'.$companyinfo->address
		);
		// $this->email->attach('application/modules/purchase/assets/excel/'.$fileName);
		$this->email->attach(FCPATH .'application/modules/purchase/assets/excel/'.$fileName);
		$this->email->send();