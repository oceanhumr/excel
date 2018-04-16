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
    public  $PHPExcel;           //一个excel对象

    public function __construct()
    {
        $this->PHPExcel=new \PHPExcel();
    }


    /**
     * 获取phpexcel实例
     * @return \PHPExcel
     */
    public function getPHPExcel()
    {
        return $this->PHPExcel;
    }


    //以下是设置excel的基本属性


    /**
     * 创建人
     * @param null $creator
     */
    public function setCreator($creator=null)
    {
        if($creator){
            $this->PHPExcel->getProperties()->setCreator($creator);
        }
        return $this;
    }


    /**
     * 最后修改人
     * @param $lastModified
     */
    public function setLastModifiedBy($lastModified=null)
    {
        if($lastModified){
            $this->PHPExcel->getProperties()->setLastModifiedBy($lastModified);
        }
        return $this;
    }


    /**
     * 标题
     * @param null $title
     */
    public function setTitle($title=null)
    {
        if($title){
            $this->PHPExcel->getProperties()->setTitle($title);
        }
        return $this;
    }


    /**
     * 题目
     * @param null $subject
     */
    public function setSubject($subject=null)
    {
        if($subject){
            $this->PHPExcel->getProperties()->setSubject($subject);
        }
        return $this;
    }

    /**
     * 描述
     * @param null $description
     */
    public function setDescription($description=null)
    {
        if($description){
            $this->PHPExcel->getProperties()->setDescription($description);
        }
        return $this;
    }

    /**
     * 关键词
     * @param null $keyWords
     */
    public function setKeywords($keyWords=null)
    {
        if($keyWords){
            $this->PHPExcel->getProperties()->setKeywords($keyWords);
        }
        return $this;
    }


    /**
     * 种类
     * @param null $category
     */
    public function setCategory($category=null)
    {
        if($category){
            $this->PHPExcel->getProperties()->setCategory($category);
        }
        return $this;
    }





    //以下是基于活动sheet的操作


    /**
     * 设置行高
     * @param $row
     * @param $height
     */
    public function setRowHeight($row,$height)
    {
        if($row&&$height){
            $this->PHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight($height);
        }
    }


    /**
     * 设置所有单元格默认的行高
     * @param $height
     */
    public function setDefaultRowHeight($height)
    {
        if($height){
            $this->PHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight($height);
        }
        return $this;
    }


    /**
     * 设置列宽
     * @param $colum
     * @param $width
     */
    public function setColumWidth($colum,$width)
    {
        if($colum&&$width){
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
        if($width){
            $this->PHPExcel->getActiveSheet()->getDefaultColumnDimension()->setWidth($width);
        }
        return $this;
    }





    /**
     * 保存成excel5的格式
     * @param null $fileName
     */
    public function saveXls($fileName=null)
    {
        // Redirect output to a client’s web browser (Excel5)
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$fileName.'.xls"');
        header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0

        $objWriter = \PHPExcel_IOFactory::createWriter($this->PHPExcel, 'Excel5');
        $objWriter->save('php://output');
        exit;
    }


    /**
     * 保存成excel7的格式
     * @param null $fileName
     */
    public function saveXlsx($fileName=null)
    {
        // Redirect output to a client’s web browser (Excel2007)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="'.$fileName.'.xlsx"');
        header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0

        $objWriter = \PHPExcel_IOFactory::createWriter($this->PHPExcel, 'Excel2007');
        $objWriter->save('php://output');
        exit;
    }

}