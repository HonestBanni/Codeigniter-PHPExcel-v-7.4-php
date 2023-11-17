<?php
defined('BASEPATH') OR exit('No direct script access allowed');


class ExcelController extends CI_Controller {

   public function excel_example(){
        
        $this->load->library('excel');
        
        $this->excel->setActiveSheetIndex(0);
        //name the worksheet
        $this->excel->getActiveSheet()->setTitle('Enter Title');
        //set cell A1 content with some text

        $this->excel->getActiveSheet()->setCellValue('A1', 'Student Name');
        $this->excel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
        $this->excel->getActiveSheet()->getStyle('A1')->getFont()->setSize(14);

        $this->excel->getActiveSheet()->setCellValue('B1', 'Class');
        $this->excel->getActiveSheet()->getStyle('B1')->getFont()->setBold(true);
        $this->excel->getActiveSheet()->getStyle('B1')->getFont()->setSize(14);
         
        $this->excel->getActiveSheet()->setCellValue('C1', 'Section');
        $this->excel->getActiveSheet()->getStyle('C1')->getFont()->setBold(true);
        $this->excel->getActiveSheet()->getStyle('C1')->getFont()->setSize(14);
         
        $this->excel->getActiveSheet()->setCellValue('D1', 'Gender');
        $this->excel->getActiveSheet()->getStyle('D1')->getFont()->setBold(true);
        $this->excel->getActiveSheet()->getStyle('D1')->getFont()->setSize(14);
         
         
        for($col = ord('A'); $col <= ord('D'); $col++):
            //set column dimension
            $this->excel->getActiveSheet()->getColumnDimension(chr($col))->setAutoSize(true);
             //change the font size
            $this->excel->getActiveSheet()->getStyle(chr($col))->getFont()->setSize(12);
            $this->excel->getActiveSheet()->getStyle(chr($col))->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
        endfor;

                      $this->db->select('student_name,class,Section,Gender'); //Selected Fields
        $db_query   = $this->db->get('student_record')->result_array(); //db query with student_record table

        $exceldata = [];
        foreach ($db_query as $row):
            $exceldata[] = $row;
        endforeach;

        $this->excel->getActiveSheet()->fromArray($exceldata, null, 'A2');        
        $filename='student_recor_'.date('d-M-Y H:i:s').'.xls'; //File Name
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$filename.'"');
        header('Cache-Control: max-age=0'); 
        $objWriter = PHPExcel_IOFactory::createWriter($this->excel, 'Excel5');  
        $objWriter->save('php://output');


    }

}



?>