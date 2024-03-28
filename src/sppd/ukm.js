import * as ExcelJS from 'exceljs'
import excel from '../assets/sppd.xlsx'
import dayjs from 'dayjs'
import { getResponseAsBuffer } from './buffer'
import 'dayjs/locale/id'
import * as style from './ukm.border'

dayjs.locale('id')

/**
 *
 * @param {object} dates
 * @param {object} data
 */
async function save(dates, { nama, nip, golongan, jabatan, tujuan, alamat }) {
    if (
        null === dates._rawValue||
        undefined === dates._rawValue ||
        0 === dates._rawValue.length
    ) {
        return
    }

    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(await getResponseAsBuffer(excel))

    let index = 1
    for (const date of dates._rawValue) {
        await createWorksheet(workbook, date, {
            nama,
            nip,
            golongan,
            jabatan,
            tujuan,
            alamat,
            nomor: index++
        })
    }

    // flush
    workbook.removeWorksheet(workbook.getWorksheet('surat_tugas').id)
    workbook.removeWorksheet(workbook.getWorksheet('sppd_depan').id)
    workbook.removeWorksheet(workbook.getWorksheet('sppd_belakang').id)

    const buffer = await workbook.xlsx.writeBuffer()
    const blob   = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    })

    const link    = document.createElement('a')
    link.href     = window.URL.createObjectURL(blob)
    link.download = `[sppd] ${tujuan} - ${snakeCase(nama[1])}`
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
}

/**
 * Create worksheets for each date provided
 *
 * @param {ExcelJS.Workbook} workbook
 * @param {Date} tanggal
 * @param {object} data
 */
async function createWorksheet(
    workbook,
    tanggal,
    { nama, nip, golongan, jabatan, tujuan, alamat, nomor }
) {
    const short_date = dayjs(tanggal).format('D-M-YYYY')
    const date       = dayjs(tanggal).format('DD MMMM YYYY')

    createDataWorksheet(workbook, { nama, nip, golongan, jabatan, tujuan, alamat, nomor })
    createSuratTugasWorksheet(workbook, short_date, date, nomor, nama[2])
    createSppdDepanWorksheet(workbook, short_date, date, nomor)
    createSppdBelakangWorksheet(workbook, short_date, date, nomor)
}

/**
 * Create data worksheet with provided data
 *
 * @param {ExcelJS.Workbook} workbook
 * @param {object} data
 * @param {int} nomor
 */
function createDataWorksheet(workbook, { nama, nip, golongan, jabatan, tujuan, alamat, nomor }) {
    const ws = workbook.getWorksheet('data')
    ws.getCell('B1').value = nama[1] ?? ''
    ws.getCell('B2').value = nip[1] ?? ''
    ws.getCell('B3').value = golongan[1] ?? ''
    ws.getCell('B4').value = jabatan[1] ?? ''
    ws.getCell('B6').value = tujuan ?? ''
    ws.getCell('B7').value = alamat ?? ''
    ws.getCell(`B${7 + nomor}`).value = nomor
    ws.getCell('C1').value = nama[2] ?? ''
    ws.getCell('C2').value = nip[2] ?? ''
    ws.getCell('C3').value = golongan[2] ?? ''
    ws.getCell('C4').value = jabatan[2] ?? ''
}

/**
 * Create Surat Tugas worksheet
 *
 * @param {ExcelJS.Workbook} workbook
 * @param {string} short_date
 * @param {string} date
 * @param {int} nomor
 * @param {string} secondName
 */
function createSuratTugasWorksheet(workbook, short_date, date, nomor, secondName) {
    const surat_tugas = workbook.getWorksheet('surat_tugas')
    let cp_surat_tugas = workbook.addWorksheet('cp')

    cp_surat_tugas.model = Object.assign(surat_tugas.model, {
        mergeCells: surat_tugas.model.merges,
    })
    cp_surat_tugas.name = `surat_tugas ${short_date}`
    cp_surat_tugas.getCell('B7').value = { formula: `"Nomor : 449.1/"& (data!B${7 + nomor}) &"/UKM/2024"`, result: 'Nomor : 449.1/ /UKM/2024' }
    cp_surat_tugas.getCell('E23').value = date
    cp_surat_tugas.getCell('E23').numFmt = '[$-id-ID]dd mmmm yyyy;@'
    // remove unsed cell
    if (undefined === secondName) {
        cp_surat_tugas.getCell('C40').value = ''
        cp_surat_tugas.getCell('C41').value = ''
        cp_surat_tugas.getCell('D40').value = ''
        cp_surat_tugas.getCell('D41').value = ''
        cp_surat_tugas.getCell('F41').value = ''
    }
}

/**
 * Create SPPD Depan worksheet
 *
 * @param {ExcelJS.Workbook} workbook
 * @param {string} short_date
 * @param {string} date
 * @param {int} nomor
 */
function createSppdDepanWorksheet(workbook, short_date, date, nomor) {
    const sppd_depan = workbook.getWorksheet('sppd_depan')
    let cp_sppd_depan = workbook.addWorksheet('cp')

    cp_sppd_depan.model = Object.assign(sppd_depan.model, {
        mergeCells: sppd_depan.model.merges,
    })
    cp_sppd_depan.name = `sppd_depan ${short_date}`
    cp_sppd_depan.getCell('A8').value = { formula: `"Nomor : 449.1/"& (data!B${7 + nomor}) &"/UKM/2024"`, result: 'Nomor : 449.1/ /UKM/2024' }
    cp_sppd_depan.getCell('D22').value = date
    cp_sppd_depan.getCell('D22').numFmt = '[$-id-ID]dd mmmm yyyy;@'
    // fix boder not render as well
    style.styleSppdDepanWorksheet(cp_sppd_depan)
}

/**
 * Create SPPD Belakang worksheet
 *
 * @param {ExcelJS.Workbook} workbook
 * @param {string} short_date
 * @param {string} date
 * @param {int} nomor
 */
function createSppdBelakangWorksheet(workbook, short_date, date, nomor) {
    const sppd_belakang = workbook.getWorksheet('sppd_belakang')
    let cp_sppd_belakang = workbook.addWorksheet('cp')

    cp_sppd_belakang.model = Object.assign(sppd_belakang.model, {
        mergeCells: sppd_belakang.model.merges,
    })
    cp_sppd_belakang.name = `sppd_belakang ${short_date}`
    cp_sppd_belakang.getCell('F5').value  = { formula: `": 449.1/"& (data!B${7 + nomor}) &"/UKM/2024"`, result: ': 449.1/ /UKM/2024' }
    cp_sppd_belakang.getCell('F7').value  = ': ' + date
    cp_sppd_belakang.getCell('D13').value = date
    cp_sppd_belakang.getCell('F7').numFmt = ': [$-id-ID]dd mmmm yyyy;@'
    // fix boder not render as well
    style.styleSppdBelakangWorksheet(cp_sppd_belakang)
}

/**
 * @param {string} title
 *
 * @return string
 */
function snakeCase(title) {
    return title
        .replace(/[.,]/g, '')
        .replace(/\W+/g, ' ')
        .split(/ |\B(?=[A-Z])/)
        .map(word => word.toLowerCase())
        .join('_')
}

export { save, createWorksheet }
