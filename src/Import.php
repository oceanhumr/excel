<?php
/**
 * Created by PhpStorm.
 * User: Hu
 * Date: 2018/4/17
 * Time: 13:47
 */

namespace ocean\excel;


class Import
{
    public $PHPExcel;   //读取文件的excel实例

    private $_file_path;//要载入的文件地址

    private $_data;     //存放execl中的数据

    /**
     * Import constructor.
     * @param $filePath 要读取的excel文件路径
     * @throws \PHPExcel_Reader_Exception
     */
    public function __construct($filePath)
    {
        $filePath=iconv('utf-8','gb2312',$filePath);
        if(is_file($filePath)){
//            $this->_file_path=iconv('utf-8','gb2312',$filePath);
            $this->_file_path=$filePath;
            $this->PHPExcel=\PHPExcel_IOFactory::load($this->_file_path);
        }
    }


    /**
     * 获取excel全部数据（下标以sheet的title为key）
     * @return mixed
     */
    public function getData()
    {
        foreach($this->PHPExcel->getWorksheetIterator() as $sheet){
            $title=$sheet->getTitle();
            $this->_data[$title]=$sheet->toArray();
        }
        return $this->_data;
    }

    
    
}