package com.cc

import org.apache.spark.sql.{DataFrame, SparkSession}

import com.crealytics.spark.excel._
import org.apache.spark.sql.types.{DoubleType, IntegerType, StructField, StructType}
import java.io.File

/**
 * Created with IntelliJ IDEA.
 * Author: CC
 * E-mail: 203717588@qq.com
 * Date: 2024/7/6
 * Time: 16:42
 * Description:
 */
object Spark2Excel {


  def main(args: Array[String]): Unit = {

    // SparkSQL
    // enableHiveSupport()：开启Hive支持需要添加依赖spark-hive_2.12，否则报错：
    // Unable to instantiate SparkSession with Hive support because Hive classes are not found.
    val spark = SparkSession.builder().master("local[*]").appName("Test")
      .enableHiveSupport().getOrCreate()

    val path = "E:\\temp\\spark\\西安大数据.xlsx"

    // 1、POI方式

    // 读取Excel内容
    val sheet = POIUtils.excelInit(path, "sheet1", true)
    val text: String = POIUtils.excelReader(sheet)
    // 转换为CSV作为临时文件
    val tmp_file = "temp.csv"
    POIUtils.excelWriter(text, tmp_file, true)

    // 读取CSV文件
    val df = spark.read
      .option("header", "true")   // 第一行是否为表头
      .option("delimiter", "\001")   // 指定列字段间的分割符
      .csv(tmp_file)

    // 打印
    df.show(10)

    // 删除临时文件
    val file = new File(tmp_file)
    file.delete()


    // 2、Spark-Excel方式

    // 可自定义数据结构（类型）
    val schema = StructType(List(
      StructField("col1", IntegerType),
      StructField("col2", DoubleType)))

    // 方式1
    val df: DataFrame = spark.read
      .format("com.crealytics.spark.excel")
      //.option("dataAddress", "Sheet2!A1")   // 指定Sheet和Cell范围，默认Sheet1!A1（第一个Sheet），指定其它Sheet需要写!A1才有效
      .option("header", "true")   // 必选参数，第一行是否为表头，未指定报错
      .option("treatEmptyValuesAsNulls", "true")   // 可选参数，是否将空单元格设置为null，默认true
      .option("inferSchema", "true")   // 可选参数，是否为推断模式，默认false
      .option("timestampFormat", "yyyy-MM-dd HH:mm:ss")   // 可选参数，格式化日期，默认yyyy-MM-dd HH:mm:ss[.fffffffff]
      .option("maxRowsInMemory", 20)   // 可选参数，默认为none，如果设置，则使用流式读取器，它可以帮助处理大文件或xls格式失败
      //.schema(CustomSchema)   // 可选参数，默认为推断模式（inferSchema=true）或所有列都是String
      .load(path)

//    // 方式2
//    val df: DataFrame = spark.read.excel(
//      //dataAddress = "Sheet2!A1",
//      header = true,
//      treatEmptyValuesAsNulls = true,
//      inferSchema = true,
//      timestampFormat = "yyyy-MM-dd HH:mm:ss",
//      maxRowsInMemory = 20
//    )
//      //.schema(CustomSchema)
//      .load(path)

    // 删除全部为NA/null值的行（默认any，删除包含NA/null值的行）
    val df1: DataFrame = df.na.drop("all")
    // 查看字段类型
    df1.dtypes.foreach(println)
    df1.printSchema()
    // 查看结果
    df1.show(10)

    // 创建临时视图
    df1.createTempView("v1")

    // SQL分析
    val sql =
      """
        |select * from v1 limit 10
        |""".stripMargin
    val data: DataFrame = spark.sql(sql)
    data.show()


  }


}
