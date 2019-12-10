package paulEMG

import koma.*
import koma.extensions.get
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream
import kotlin.math.absoluteValue
import kotlin.math.pow
import kotlin.reflect.KFunction


// 2D matrix for a specific posture i.e. spher_ch1
typealias DataMatrix = List<List<Double>>

//sheet name -> 2D Matrix
typealias EMGSignal = Map<String, DataMatrix>

//MAV, CC...etc function
typealias EigenFunction = KFunction<EMGSignal>

//(MAV, CC) , (MAV, WAMP)...etc
typealias EigenMethodPair = Pair<EigenFunction, EigenFunction>

//SD for method pair (MAV, CC)
data class EigenvaluesSD(val lambdaSD: Double, val detSD: Double, val traceSD: Double)

//((MAV, CC), (Lambda, determinant, trace)
typealias EigenMethodPairSD = Pair<EigenMethodPair, EigenvaluesSD>

//standard deviation
fun DoubleArray.sd(): Double {
    val mean = average()
    val sd = fold(0.0, { accumulator, next -> accumulator + (next - mean).pow(2.0) })
    return kotlin.math.sqrt(sd / size)
}

//for iteration purpose
val functionList = listOf(::MAV, ::CC, ::BZC, ::LSG, ::WAMP, ::SSC)

//chunck size for function
val chunkWidth = 30

val function_combination_list = ArrayList<EigenMethodPair>().apply {

    //till second last
    for (index in 0 until functionList.size - 1) {
        for (index_2 in index + 1 until functionList.size)
            add(EigenMethodPair(functionList[index], functionList[index_2]))
    }
}

val mine_EMG = "./mine_EMG.xlsx"
val female_1 = "./female_1.xlsx"
val female_2 = "./female_2.xlsx"
val female_3 = "./female_3.xlsx"
val male_1 = "./male_1.xlsx"

fun main() {

    //get mine first
    val excel_mine_EMG = exportStandardDeviationFile(mine_EMG)
    exportExcel("./sd_mine_EMG.xlsx", excel_mine_EMG)

    val excel_male_1 = exportStandardDeviationFile(male_1)
    exportExcel("./sd_male_1.xlsx", excel_male_1)

    val excel_female_1 = exportStandardDeviationFile(female_1)
    exportExcel("./sd_female_1.xlsx", excel_female_1)

    val excel_female_2 = exportStandardDeviationFile(female_2)
    exportExcel("./sd_female_2.xlsx", excel_female_2)

    val excel_female_3 = exportStandardDeviationFile(female_3)
    exportExcel("./sd_female_3.xlsx", excel_female_3)

    println()

}

fun coefficientFunction(SD_1: EigenvaluesSD, SD_2: EigenvaluesSD): Int {
    val SD_1_result = SD_1.lambdaSD <= 2.0 && SD_1.detSD <= 0.5 && SD_1.traceSD <= 3
    val SD_2_result = SD_2.lambdaSD <= 2.0 && SD_2.detSD <= 0.5 && SD_2.traceSD <= 3

    return if (SD_1_result && SD_2_result) 1 else 0
}

fun findMatchingPairSD(mine_pairSD: EigenvaluesSD, dataSet_map: Map<String, List<EigenMethodPairSD>>) {

    val dataSet_map_ch1 = dataSet_map.filter {
        it.key.contains("ch1")
    }

    val matching_gesture = Pair("", 0.0)

    for((gesture, eigenValue) in dataSet_map_ch1){

    }

}

//fun gestureMatchingValue(mine_gesture : EigenvaluesSD, comparingGesture : EigenvaluesSD) : Double{
//
//}

fun exportExcel(filePath: String, map: Map<String, List<EigenMethodPairSD>>) {

    // Create a Workbook
    val workbook: Workbook = XSSFWorkbook() // new HSSFWorkbook() for generating `.xls` file

    for ((sheetName, listOfPairSD) in map) {

        val sheet = workbook.createSheet(sheetName)
        val headerRow = sheet.createRow(0)
        val lambdaRow = sheet.createRow(1)
        val detRow = sheet.createRow(2)
        val traceRow = sheet.createRow(3)
        val averageRow = sheet.createRow(4)


        // Create cells
        for (i in listOfPairSD.indices) {

            val function_name = listOfPairSD[i].first.first.name + " , " + listOfPairSD[i].first.second.name
            headerRow.createCell(i * 3 + 1).setCellValue(function_name)

            lambdaRow.createCell(i * 3 + 1).setCellValue(listOfPairSD[i].second.lambdaSD)
            detRow.createCell(i * 3 + 1).setCellValue(listOfPairSD[i].second.detSD)
            traceRow.createCell(i * 3 + 1).setCellValue(listOfPairSD[i].second.traceSD)

            averageRow.createCell(i * 3 + 1).setCellValue(
                (listOfPairSD[i].second.lambdaSD + listOfPairSD[i].second.traceSD
                        + listOfPairSD[i].second.detSD) / 3.0
            )

        }

        lambdaRow.createCell(0).setCellValue("Lambda SD")
        detRow.createCell(0).setCellValue("Det SD")
        traceRow.createCell(0).setCellValue("Trace SD")
        averageRow.createCell(0).setCellValue("Average")

    }

    // Write the output to a file
    // Write the output to a file
    val fileOut = FileOutputStream(filePath.substring(2))
    workbook.write(fileOut)
    fileOut.close()

    // Closing the workbook
    // Closing the workbook
    workbook.close()


}

fun exportStandardDeviationFile(filePath: String): Map<String, List<EigenMethodPairSD>> {

    val map = generateMapWithFile(filePath)

    val file_SD_ValueForEach = HashMap<String, List<EigenMethodPairSD>>()

    val function_map =
        HashMap<KFunction<EMGSignal>,
                EMGSignal>().apply {
            put(::MAV, MAV(map, chunkWidth))
            put(::CC, CC(map, chunkWidth))
            put(::BZC, BZC(map, chunkWidth))
            put(::LSG, LSG(map, chunkWidth))
            put(::WAMP, WAMP(map, chunkWidth))
            put(::SSC, SSC(map, chunkWidth))
        }

    for (sheetName in map.keys) {

        val listOfEigenvaluesSD = ArrayList<EigenMethodPairSD>()

        for (eigenMethodPair in function_combination_list) {
            val (method_1, method_2) = eigenMethodPair
            val dataMatrix_1 = function_map[method_1]!![sheetName]!!
            val dataMatrix_2 = function_map[method_2]!![sheetName]!!

            val eigenvaluesSD = getEigenvaluesSDForTwoDataMatrices(dataMatrix_1, dataMatrix_2)
            listOfEigenvaluesSD.add(EigenMethodPairSD(eigenMethodPair, eigenvaluesSD))
        }

        file_SD_ValueForEach[sheetName] = listOfEigenvaluesSD

    }

    return file_SD_ValueForEach
}

fun getEigenvaluesSDForTwoDataMatrices(
    dataMatrix_1: DataMatrix,
    dataMatrix_2: DataMatrix
): EigenvaluesSD {

    val lambda_list = ArrayList<Double>()
    val det_list = ArrayList<Double>()
    val trace_list = ArrayList<Double>()

    for (index in dataMatrix_1.indices) {

        val dataMatrix_1_Row = dataMatrix_1[index]
        val dataMatrix_2_Row = dataMatrix_2[index]

        val (lambda, det) = findLambdaOf_Two_1D_Matrix(dataMatrix_1_Row, dataMatrix_2_Row)

        val original_4D_matrix = create(
            arrayOf(
                dataMatrix_1_Row.toDoubleArray(), dataMatrix_2_Row.toDoubleArray()

            )
        )

        val original_trace = original_4D_matrix.trace()

        lambda_list.add(lambda)
        det_list.add(det)
        trace_list.add(original_trace)
    }

    return EigenvaluesSD(
        lambda_list.toDoubleArray().sd(),
        det_list.toDoubleArray().sd(),
        trace_list.toDoubleArray().sd()
    )
}


//return lambda of two matrix and C's det
fun findLambdaOf_Two_1D_Matrix(m1: List<Double>, m2: List<Double>): Pair<Double, Double> {

    //get average
    val m1_mean = m1.average()
    val m2_mean = m2.average()

    //minus each element in the arraylist by the mean
    val m1_prime = m1.toMutableList()
    val m2_prime = m2.toMutableList()

    m1_prime.forEachIndexed { index, d ->
        m1_prime[index] = d - m1_mean
    }

    m2_prime.forEachIndexed { index, d ->
        m2_prime[index] = d - m2_mean
    }

    val matrix_2D_prime = create(arrayOf(m1_prime.toDoubleArray(), m2_prime.toDoubleArray()))
    val matrix_2D_prime_transpose = matrix_2D_prime.transpose()

    val C = matrix_2D_prime.times(matrix_2D_prime_transpose).times(0.01)

    val i = C[0, 0]
    val j = C[0, 1]
    val k = C[1, 1]

    val lambda_1 = 0.5 * ((i + k) + sqrt(pow(i + k, 2) - 4 * (i * k - pow(j, 2))))
    val lambda_2 = 0.5 * ((i + k) - sqrt(pow(i + k, 2) - 4 * (i * k - pow(j, 2))))

    return Pair(max(lambda_1, lambda_2), C.det())
}

fun CC(
    map: EMGSignal,
    w: Int
): EMGSignal {

    val function: (chuckedList: List<Double>) -> Double = { chunk ->

        var accumulator = 0.0

        for (index in 1 until chunk.size) {
            accumulator += (chunk[index].absoluteValue - chunk[index - 1].absoluteValue).absoluteValue
        }

        accumulator
    }

    return getRetMap(map, w, function)

}

fun BZC(
    map: EMGSignal,
    w: Int
): EMGSignal {

    val function: (chuckedList: List<Double>) -> Double = { chunk ->

        val bias = chunk.average()

        var accumulator = 0.0

        for (index in 1 until chunk.size) {
            accumulator +=
                if (((chunk[index] - bias) * (chunk[index - 1] - bias)) > 0)
                    1
                else
                    0
        }

        accumulator

    }

    return getRetMap(map, w, function)

}

fun SSC(
    map: EMGSignal,
    w: Int
): EMGSignal {

    val function: (chuckedList: List<Double>) -> Double = { chunk ->

        val threshold = chunk.average()

        var accumulator = 0.0

        for (index in 2 until chunk.size) {

            val si = chunk[index] - chunk[index - 1]
            val si_1 = chunk[index - 1] - chunk[index - 2]

            accumulator +=
                if (si * si_1 > threshold)
                    1
                else
                    0
        }

        accumulator

    }

    return getRetMap(map, w, function)

}

fun LSG(
    map: EMGSignal,
    w: Int
): EMGSignal {

    val mavMap = MAV(map, w)

    for (value in mavMap.values) {

        value.forEach {
            for (index in it.size - 1 downTo 1) {

                val replacement = it[index] - it[index - 1]
                (it as ArrayList)[index] = replacement
            }
        }
    }

    return mavMap
}

fun WAMP(
    map: EMGSignal,
    w: Int
): EMGSignal {

    val function: (chuckedList: List<Double>) -> Double = { chunk ->

        val threshold = chunk.average()

        var accumulator = 0.0

        for (index in 1 until chunk.size) {
            accumulator +=
                if ((chunk[index] - chunk[index - 1]).absoluteValue > threshold)
                    1
                else
                    0
        }

        accumulator

    }

    return getRetMap(map, w, function)

}

fun WL(map: EMGSignal, w: Int): EMGSignal {

    val function: (chuckedList: List<Double>) -> Double = { chunk ->

        var accumulator = 0.0

        for (index in 1 until chunk.size) {
            accumulator += (chunk[index] - chunk[index - 1]).absoluteValue
        }

        accumulator
    }

    return getRetMap(map, w, function)
}

fun MAV(map: EMGSignal, w: Int): EMGSignal {

    val function: (chuckedList: List<Double>) -> Double = { chunk ->
        chunk.sumByDouble { it.absoluteValue } / chunk.size
    }

    return getRetMap(map, w, function)
}

fun getRetMap(
    map: EMGSignal,
    w: Int,
    function: (chuckedList: List<Double>) -> Double
): EMGSignal {

    //returned Map for result of the same type but applied logic function to chunked List
    val retMap = HashMap<String, DataMatrix>()

    //sheet to value matrix
    for ((sheet, value) in map) {

        val retLists = ArrayList<List<Double>>()

        //each row in 2D value matrix
        for (row in value) {

            val cells = ArrayList<Double>()

            //chunk 1D row value matrix to (/w) chunks
            val chunks = row.chunked(w)

            //apply function to each chunks and add to return cells
            for (each in chunks) {
                cells.add(function(each))
            }

            retLists.add(cells)
        }

        retMap[sheet] = retLists
    }

    return retMap
}


fun generateMapWithFile(filePath: String): EMGSignal {

    val map = HashMap<String, DataMatrix>()

    val inputStream = FileInputStream(filePath)
    //Instantiate Excel workbook using existing file:
    val workBook = WorkbookFactory.create(inputStream)

    for (i in 0 until workBook.numberOfSheets) {

        val sheet = workBook.getSheetAt(i)
        val rowList = ArrayList<List<Double>>()

        for (row in 0 until if (filePath == mine_EMG) sheet.physicalNumberOfRows else 30) {

            val columnList = ArrayList<Double>()

            for (column in 0 until 3000) {
                columnList.add(sheet.getRow(row).getCell(column).numericCellValue)
            }

            rowList.add(columnList)
        }

        map[sheet.sheetName] = rowList
    }

    return map
}


