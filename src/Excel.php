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
    public static $PHPExcel;           //一个excel对象

    public function __construct()
    {


    }


    public function getPHPExcel()
    {
        return self::$PHPExcel;
    }


    //以下是设置excel的基本属性


    /**
     * 创建人
     * @param null $creator
     */
    public function setCreator($creator=null)
    {
        
    }


    /**
     * 最后修改人
     * @param $lastModified
     */
    public function setLastModifiedBy($lastModified=null)
    {
        
    }


    /**
     * 标题
     * @param null $title
     */
    public function setTitle($title=null)
    {
        
    }


    /**
     * 题目
     * @param null $subject
     */
    public function setSubject($subject=null)
    {
        
    }

    /**
     * 描述
     * @param null $description
     */
    public function setDescription($description=null)
    {
        
    }

    /**
     * 关键词
     * @param null $keyWords
     */
    public function setKeywords($keyWords=null)
    {
        
    }


    /**
     * 种类
     * @param null $category
     */
    public function setCategory($category=null)
    {
        
    }

}