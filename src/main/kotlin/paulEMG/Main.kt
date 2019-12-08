package paulEMG

import koma.*
import koma.extensions.get
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.FileInputStream
import kotlin.math.absoluteValue

fun main() {

    val mine_EMG = "./mine_EMG.xlsx"
    val female_1 = "./female_1.xlsx"
    val female_2 = "./female_2.xlsx"
    val female_3 = "./female_3.xlsx"
    val male_1 = "./male_1.xlsx"

//    // Create some normal random noise
//    var a = randn(100, 2)
//    var b = cumsum(a)
//
//    figure(1)
//    // Second parameter is color
//    plot(a, 'b', "First Run")
//    plot(a, 'y', "First Run Offset")
//    xlabel("Time (s)")
//    ylabel("Magnitude")
//    title("White Noise")
//
//    figure(2)
//    plot(b, 'g') // green
//    xlabel("Velocity (lightweeks/minute)")
//    ylabel("Intelligence")
//    title("Random Walk")


    val map = generateMapWithFile(male_1)


    val mavMap = MAV(map, 30)
    val ccMap = CC(map, 30)
    val wampMap = WAMP(map, 30)
    val bzcpMap = BZC(map, 30)
    val sscMap = SSC(map, 30)
    val lsgpMap = LSG(map, 30)

    val a = mavMap["cyl_ch1"]!![0]
    val b = ccMap["cyl_ch1"]!![0]

    val x = findLambdaOf_Two_1D_Matrix(a, b)

    println()

}

fun findLambdaOf_Two_1D_Matrix(m1: List<Double>, m2: List<Double>): Double {

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

    return max(lambda_1, lambda_2)
}

fun CC(
    map: Map<String, List<List<Double>>>,
    w: Int
): Map<String, List<List<Double>>> {

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
    map: Map<String, List<List<Double>>>,
    w: Int
): Map<String, List<List<Double>>> {

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
    map: Map<String, List<List<Double>>>,
    w: Int
): Map<String, List<List<Double>>> {

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
    map: Map<String, List<List<Double>>>,
    w: Int
): Map<String, List<List<Double>>> {

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
    map: Map<String, List<List<Double>>>,
    w: Int
): Map<String, List<List<Double>>> {

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

fun WL(map: Map<String, List<List<Double>>>, w: Int): Map<String, List<List<Double>>> {

    val function: (chuckedList: List<Double>) -> Double = { chunk ->

        var accumulator = 0.0

        for (index in 1 until chunk.size) {
            accumulator += (chunk[index] - chunk[index - 1]).absoluteValue
        }

        accumulator
    }

    return getRetMap(map, w, function)
}

fun MAV(map: Map<String, List<List<Double>>>, w: Int): Map<String, List<List<Double>>> {

    val function: (chuckedList: List<Double>) -> Double = { chunk ->
        chunk.sumByDouble { it.absoluteValue } / chunk.size
    }

    return getRetMap(map, w, function)
}

fun getRetMap(
    map: Map<String, List<List<Double>>>,
    w: Int,
    function: (chuckedList: List<Double>) -> Double
): Map<String, List<List<Double>>> {

    //returned Map for result of the same type but applied logic function to chunked List
    val retMap = HashMap<String, List<List<Double>>>()

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


fun generateMapWithFile(filePath: String): Map<String, List<List<Double>>> {

    val map = HashMap<String, List<List<Double>>>()

    val inputStream = FileInputStream(filePath)
    //Instantiate Excel workbook using existing file:
    val workBook = WorkbookFactory.create(inputStream)

    for (i in 0 until workBook.numberOfSheets) {

        val sheet = workBook.getSheetAt(i)
        val rowList = ArrayList<List<Double>>()

        for (row in 0 until sheet.physicalNumberOfRows) {

            val columnList = ArrayList<Double>()

            for (column in 0 until sheet.getRow(row).physicalNumberOfCells) {
                columnList.add(sheet.getRow(row).getCell(column).numericCellValue)
            }

            rowList.add(columnList)
        }

        map[sheet.sheetName] = rowList
    }

    return map
}


