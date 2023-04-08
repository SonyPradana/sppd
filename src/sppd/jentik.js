import * as ExcelJS from 'exceljs'
import excel from '../assets/sppd.xlsx'
import '@vuepic/vue-datepicker/dist/main.css'
import dayjs from 'dayjs'
import { getResponseAsBuffer } from './buffer'

/**
 * 
 * @param {object} dates 
 * @param {object} data 
 */
async function save(dates, data) {
  if (dates._rawValue === null) {
    return
  }
  
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(await getResponseAsBuffer(excel))

  dates._rawValue.forEach(async date => {
    await create(workbook, date, data)
  })

  // flush
  workbook.removeWorksheet(workbook.getWorksheet('surat_tugas').id);
  workbook.removeWorksheet(workbook.getWorksheet('sppd_depan').id);
  workbook.removeWorksheet(workbook.getWorksheet('sppd_belakang').id);
  
  const buffer = await workbook.xlsx.writeBuffer()
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  })
  
  const link = document.createElement('a')
  link.href = window.URL.createObjectURL(blob)
  link.download = `sppd ${data.nama}`
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
}

/**
 * Create sheat from workbook
 * 
 * @param {ExcelJS.Workbook} workbook 
 * @param {Date} date 
 * @param {object} data 
 */
async function create(workbook, date, data) {
  // filen name
  const day = date.getDate()
  const month = date.getMonth() + 1
  const year = date.getFullYear()

  const ws = workbook.getWorksheet('data')
  ws.getCell('B1').value = data.nama
  ws.getCell('B2').value = data.nip
  ws.getCell('B3').value = data.golongan
  ws.getCell('B4').value = data.jabatan
  // ws.getCell('B5').value = date
  ws.getCell('B6').value = data.tujuan
  ws.getCell('B7').value = data.alamat

  // cp surat tugas
  const surat_tugas = workbook.getWorksheet('surat_tugas')
  let cp_surat_tugas = workbook.addWorksheet('cp')

  cp_surat_tugas.model = Object.assign(surat_tugas.model, {
    mergeCells: surat_tugas.model.merges
  })
  cp_surat_tugas.name = `surat_tugas ${day}-${month}-${year}`
  cp_surat_tugas.getCell('E23').value = dayjs(date).format('DD MMMM YYYY')
  cp_surat_tugas.getCell('E23').numFmt = '[$-id-ID]dd mmmm yyyy@'

  // cp sppd_depan
  const sppd_depan = workbook.getWorksheet('sppd_depan')
  let cp_sppd_depan = workbook.addWorksheet('cp')

  cp_sppd_depan.model = Object.assign(sppd_depan.model, {
    mergeCells: sppd_depan.model.merges
  })
  cp_sppd_depan.name = `sppd_depan ${day}-${month}-${year}`
  cp_sppd_depan.getCell('D22').value = dayjs(date).format('DD MMMM YYYY')
  cp_sppd_depan.getCell('D22').numFmt = '[$-id-ID]dd mmmm yyyy@'

  // cp sppd_belakang
  const sppd_belakang = workbook.getWorksheet('sppd_belakang')
  let cp_sppd_belakang = workbook.addWorksheet('cp')

  cp_sppd_belakang.model = Object.assign(sppd_belakang.model, {
    mergeCells: sppd_belakang.model.merges
  });
  cp_sppd_belakang.name = `sppd_belakang ${day}-${month}-${year}`;
  cp_sppd_belakang.getCell('F7').value = dayjs(date).format('DD MMMM YYYY')
  cp_sppd_belakang.getCell('F7').numFmt = '[$-id-ID]dd mmmm yyyy;@'
}

export {
  save,
  create
}
