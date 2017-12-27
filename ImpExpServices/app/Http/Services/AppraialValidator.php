<?php
/**
 * Created by PhpStorm.
 * User: imake
 * Date: 12/21/17
 * Time: 11:55 PM
 */

namespace App\Http\Services;
use App\Model\AppraisalStructureModel;
use Excel;
use Log;
class AppraialValidator
{
    public  function validateTemplate($numberOfSheet,$request,$header_values,$bank_values,$number_values,$all_number_values){
        $structure_id = $request->structure_id;
        $appraisalStructure = AppraisalStructureModel::find($structure_id);
        $structure_name = $appraisalStructure->structure_name;
        //Log::info($structure_name);

        $this->result_status = 1;
        $this->code = "S001";
        $this->msg = "import Success";
        //$f='/Users/imake/Desktop/detail_import_okr.xlsx';
        foreach ($request->file() as $f) {
            for ($k = 0;$k<$numberOfSheet ; $k++) {
                //Log::info('into looop '.$k);
                Excel::selectSheetsByIndex($k)->load($f, function($reader) use ($header_values, $all_number_values, $k, $bank_values, $number_values,$structure_name) {
                    $sheet =  $reader->getExcel()->getSheet($k);
                    $sheet_name = $sheet->getTitle();
                    $pos = strpos($sheet_name,$structure_name);
                    //Log::info('pos['.$pos.']');
                    //Log::info(gettype($pos));
                    if(!strlen((string)$pos)>0){
                        $this->result_status = 0;
                        $this->code = 'E004';
                        $this->msg = "เลือก Template ไม่ถูกต้อง";
                        goto end;
                    }
                    foreach ($header_values as $header) {
                        $head_val = $sheet->getCell($header.'1')->getValue() ;
                        if ( !(!empty($head_val) && strlen(trim($head_val))>0) ) {
                            $this->result_status = 0;
                            $this->code = 'E005';
                            $this->msg = "Template ไม่ถูกต้อง Sheet[".$sheet_name.']!'.$header.'1';
                            goto end;
                        }
                    }
                    foreach ($header_values as $header) {
                        $head_val = $sheet->getCell($header.'2')->getValue() ;
                        if ( !(!empty($head_val) && strlen(trim($head_val))>0) ) {
                            $this->result_status = 0;
                            $this->code = 'E006';
                            $this->msg = "กรุณากรอกข้อมูลช่อง Sheet[".$sheet_name.']!'.$header.'2';
                            goto end;
                        }
                    }
                    $head_val = $sheet->getCell("A1")->getValue() ;
                    Log::info('head_val['.$head_val.']');
                    for ($i = 2; ; $i++) {
                        $cds_name = $sheet->getCell('A'.$i)->getValue() ;
                        if ( !empty($cds_name) && strlen(trim($cds_name))>0 ) {
                            //Log::info($cds_name);
                            foreach ($bank_values as $bank) {
                                $bank_val = $sheet->getCell($bank.$i)->getValue() ;
                                if ( !empty($bank_val) && strlen(trim($bank_val))>0 ) {
                                    $this->result_status = 0;
                                    $this->code = 'E002';
                                    $this->msg = "กรุณากรอกข้อมูลช่อง Sheet[".$sheet_name.']!'.$bank.$i;
                                    goto end;
                                }
                            }
                            foreach ($number_values as $number) {
                                $num_val = $sheet->getCell($number.$i)->getValue() ;
                                //Log::info($num_val);
                                if(!is_numeric($num_val) ){
                                    $this->result_status = 0;
                                    $this->code = 'E003';
                                    $this->msg = "กรุณากรอกข้อมูล ตัวเลข ช่อง Sheet[".$sheet_name.']!'.$number.$i;
                                    goto end;
                                }
                            }
                            foreach ($all_number_values as $all_number) {
                                $all_number_val = $sheet->getCell($all_number.$i)->getValue() ;
                                //Log::info($num_val);
                                if(strtolower(trim($all_number_val)) != 'all' && !is_numeric($all_number_val) ){
                                    $this->result_status = 0;
                                    $this->code = 'E003';
                                    $this->msg = "กรุณากรอกข้อมูล ตัวเลข ช่อง Sheet[".$sheet_name.']!'.$all_number.$i;
                                    goto end;
                                }
                            }

                        }else{
                            break;
                        }
                    }
                    end:
                });
                if($this->result_status == 0)
                    break;
            }

        }

        $result_obj = new \stdClass; // Instantiate stdClass object
        $result_obj->result_status = $this->result_status;
        $result_obj->code = $this->code;
        $result_obj->msg = $this->msg;
        return $result_obj;
    }
}