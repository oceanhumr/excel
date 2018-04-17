<?php
/**
 * Created by PhpStorm.
 * User: Hu
 * Date: 2018/4/16
 * Time: 14:46
 */

namespace ocean\excel;


class Excel
{
    public $PHPExcel;               //一个excel对象

    public $properties;             //关于excel属性操作的对象

    public $valid_colum_arr;      //有效列

    public function __construct()
    {
        $this->valid_colum_arr=range('A','Z');
        $this->PHPExcel = new \PHPExcel();
        $this->properties = $this->PHPExcel->getProperties();
    }


    /**
     * 获取phpexcel实例
     * @return \PHPExcel
     */
    public function getPHPExcel()
    {
        return $this->PHPExcel;
    }


    /**
     * 获取设置excel属性的实例
     * @return \PHPExcel_DocumentProperties
     */
    public function getProperties()
    {
        return $this->properties;
    }


    /**
     * 获取第几个sheet
     * @param $i
     */
    public function getSheet($i)
    {
        //现在已有的sheet的个数
        $count = $this->PHPExcel->getSheetCount();

        //要选择的sheet已经存在了
        if ($i >= $count) {
            for($a=$count;$a<=$i;$a++){
                $this->PHPExcel->createSheet($a);
            }
        }
        
        $this->PHPExcel->setActiveSheetIndex($i);
        return $this->PHPExcel->getActiveSheet();
    }






    //以下是基于活动sheet的操作

    /**
     * 设置行高
     * @param $row
     * @param $height
     */
    public function setRowHeight($row, $height)
    {
        if ($row && $height) {
            $this->PHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight($height);
        }
    }


    /**
     * 设置所有单元格默认的行高
     * @param $height
     */
    public function setDefaultRowHeight($height)
    {
        if ($height) {
            $this->PHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight($height);
        }
        return $this;
    }


    /**
     * 设置列宽
     * @param $colum
     * @param $width
     */
    public function setColumWidth($colum, $width)
    {
        if ($colum && $width) {
            $this->PHPExcel->getActiveSheet()->getColumnDimension($colum)->setWidth($width);
        }
        return $this;
    }


    /**
     * 设置所有单元格默认的宽
     * @param $width
     */
    public function setDefaultColumWidth($width)
    {
        if ($width) {
            $this->PHPExcel->getActiveSheet()->getDefaultColumnDimension()->setWidth($width);
        }
        return $this;
    }



    //以下导出文件的头部操作

    /**
     * 保存成excel5的格式
     * @param null $fileName
     */
    public function saveXls($fileName = null)
    {
        // Redirect output to a client’s web browser (Excel5)
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="' . $fileName . '.xls"');
        header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0

        $objWriter = \PHPExcel_IOFactory::createWriter($this->PHPExcel, 'Excel5');
        $objWriter->save('php://output');
        exit;
    }


    /**
     * 保存成excel7的格式
     * @param null $fileName
     */
    public function saveXlsx($fileName = null)
    {
        // Redirect output to a client’s web browser (Excel2007)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $fileName . '.xlsx"');
        header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0

        $objWriter = \PHPExcel_IOFactory::createWriter($this->PHPExcel, 'Excel2007');
        $objWriter->save('php://output');
        exit;
    }

}