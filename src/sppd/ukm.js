import * as ExcelJS from 'exceljs'
import excel from '../assets/sppd.xlsx'
import dayjs from 'dayjs'
import { getResponseAsBuffer } from './buffer'
import 'dayjs/locale/id'

dayjs.locale('id')

/**
 * 
 * @param {object} dates 
 * @param {object} data 
 */
async function save(dates, {nama, nip, golongan, jabatan, tujuan, alamat}) {
  if (dates._rawValue === null || dates._rawValue === undefined || dates._rawValue === []) {
    return
  }
  
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(await getResponseAsBuffer(excel))
  
  dates._rawValue.forEach(async date => {
    await create(workbook, date, {nama, nip, golongan, jabatan, tujuan, alamat})
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
  link.download = `sppd ${nama[1]}`
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
}

/**
 * Create sheat from workbook
 * 
 * @param {ExcelJS.Workbook} workbook 
 * @param {Date} tanggal 
 * @param {object} data 
 */
async function create(workbook, tanggal, {nama, nip, golongan, jabatan, tujuan, alamat}) {
  const short_date = dayjs(tanggal).format('D-M-YYYY')
  const date = dayjs(tanggal).format('DD MMMM YYYY')

  const ws = workbook.getWorksheet('data')
  ws.getCell('B1').value = nama[1] ?? ''
  ws.getCell('B2').value = nip[1] ?? ''
  ws.getCell('B3').value = golongan[1] ?? ''
  ws.getCell('B4').value = jabatan[1] ?? ''
  ws.getCell('B6').value = tujuan ?? ''
  ws.getCell('B7').value = alamat ?? ''
  // 
  ws.getCell('C1').value = nama[2] ?? ''
  ws.getCell('C2').value = nip[2] ?? ''
  ws.getCell('C3').value = golongan[2] ?? ''
  ws.getCell('C4').value = jabatan[2] ?? ''
  
  // cp surat tugas
  const surat_tugas = workbook.getWorksheet('surat_tugas')
  let cp_surat_tugas = workbook.addWorksheet('cp')

  cp_surat_tugas.model = Object.assign(surat_tugas.model, {
    mergeCells: surat_tugas.model.merges
  })
  cp_surat_tugas.name = `surat_tugas ${short_date}`
  cp_surat_tugas.getCell('E23').value   = date
  cp_surat_tugas.getCell('E23').numFmt = '[$-id-ID]dd mmmm yyyy;@'
  // remove unsed cell
  if (nama[2] === undefined) {
    cp_surat_tugas.getCell('C40').value = ''
    cp_surat_tugas.getCell('C41').value = ''
    cp_surat_tugas.getCell('D40').value = ''
    cp_surat_tugas.getCell('D41').value = ''
    cp_surat_tugas.getCell('F41').value = ''
  }

  // cp sppd_depan
  const sppd_depan = workbook.getWorksheet('sppd_depan')
  let cp_sppd_depan = workbook.addWorksheet('cp')

  cp_sppd_depan.model = Object.assign(sppd_depan.model, {
    mergeCells: sppd_depan.model.merges
  })
  cp_sppd_depan.name = `sppd_depan ${short_date}`
  cp_sppd_depan.getCell('D22').value = date
  cp_sppd_depan.getCell('D22').numFmt = '[$-id-ID]dd mmmm yyyy;@'

  // cp sppd_belakang
  const sppd_belakang = workbook.getWorksheet('sppd_belakang')
  let cp_sppd_belakang = workbook.addWorksheet('cp')

  cp_sppd_belakang.model = Object.assign(sppd_belakang.model, {
    mergeCells: sppd_belakang.model.merges
  });
  cp_sppd_belakang.name = `sppd_belakang ${short_date}`;
  cp_sppd_belakang.getCell('F7').value = date
  cp_sppd_belakang.getCell('F7').numFmt = '[$-id-ID]dd mmmm yyyy;@'
}

export {
  save,
  create
}
