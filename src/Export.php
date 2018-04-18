<?php
/**
 * Created by PhpStorm.
 * User: Hu
 * Date: 2018/4/17
 * Time: 13:39
 */

namespace ocean\excel;


class Export extends Excel
{

    private $_header;

    private $_data;

    private $_colum_number;



    /**
     * 设置excel第一行的说明文件
     * @param $arr
     */
    public function setHeader($arr)
    {
        if(is_array($arr)){
            $this->_header=$arr;
            $this->_colum_number=count($this->_header);
            //设置第一行内容的格式
            $this->setFirstRowStyle();

            return $this;
        }
    }


    /**
     * 设置默认的第一行的样式
     */
    private function setFirstRowStyle()
    {
        $activeSheet=$this->PHPExcel->getActiveSheet();
        $this->PHPExcel->getActiveSheet()->setSharedStyle($this->sharedStyle, "A1:".$this->valid_colum_arr[$this->_colum_number-1].'1');
        for($i=0;$i<$this->_colum_number;$i++){
            $activeSheet->getStyle($this->valid_colum_arr[$i].'1')->getFont()->setBold(true);
            $activeSheet->getStyle($this->valid_colum_arr[$i].'1')->getFont()->setSize(12);
            $activeSheet->getStyle($this->valid_colum_arr[$i].'1')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $activeSheet->getStyle($this->valid_colum_arr[$i].'1')->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);
        }
    }


    /**
     * 载入导出的数据
     * @param $data
     */
    public function loadData($data)
    {
        array_unshift($data,$this->_header);
        if(is_array($data)){
            $this->_data=$data;
            $this->PHPExcel->getActiveSheet()->setSharedStyle($this->sharedStyle, "A2:".$this->valid_colum_arr[$this->_colum_number-1].count($this->_data));
            $this->PHPExcel->getActiveSheet()->fromArray($data);
            return $this;
        }
    }


}