<script async setup>
import * as ExcelJS from 'exceljs'
import excel from '../../assets/sppd.xlsx'
import axios from 'axios'
import { Buffer } from 'buffer'
import { ref } from 'vue'
import VueDatePicker from '@vuepic/vue-datepicker';
import '@vuepic/vue-datepicker/dist/main.css'
import * as dayjs from 'dayjs'

const nama = ref('')
const nip = ref('')
const golongan = ref('')
const jabatan = ref('')
const tanggal = ref()
const alamat = ref('')
const tujuan = ref('')

async function getResponseAsBuffer(url) {
  try {
    const response = await axios.get(url, { responseType: 'arraybuffer' });
    const buffer = Buffer.from(response.data, 'binary');
    return buffer;
  } catch (error) {
    console.error(error);
    return null;
  }
}

async function save() {
  if (tanggal._rawValue === null) {
    return
  }
  
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(await getResponseAsBuffer(excel))

  tanggal._rawValue.forEach(async date => {
    await cread(workbook, date)
  })

  // flush
  workbook.removeWorksheet(workbook.getWorksheet('surat_tugas').id);
  workbook.removeWorksheet(workbook.getWorksheet('sppd_depan').id);
  workbook.removeWorksheet(workbook.getWorksheet('sppd_belakang').id);
  
  const buffer = await workbook.xlsx.writeBuffer()
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  
  const link = document.createElement('a')
  link.href = window.URL.createObjectURL(blob)
  link.download = `sppd ${nama.value}`
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
}

async function cread(workbook, date) {
  const formated_date = dayjs(date).format('DD MMMM YYYY')

  const ws = workbook.getWorksheet('data')
  ws.getCell('B1').value = nama.value
  ws.getCell('B2').value = nip.value
  ws.getCell('B3').value = golongan.value
  ws.getCell('B4').value = jabatan.value
  // ws.getCell('B5').value = date
  ws.getCell('B6').value = tujuan.value
  ws.getCell('B7').value = alamat.value

  // cp surat tugas
  const surat_tugas = workbook.getWorksheet('surat_tugas')
  let cp_surat_tugas = workbook.addWorksheet('cp')

  cp_surat_tugas.model = Object.assign(surat_tugas.model, {
    mergeCells: surat_tugas.model.merges
  })
  cp_surat_tugas.name = `surat_tugas ${formated_date}`
  cp_surat_tugas.getCell('E23').value = formated_date
  cp_surat_tugas.getCell('E23').numFmt = '[$-id-ID]dd mmmm yyyy@'

  // cp sppd_depan
  const sppd_depan = workbook.getWorksheet('sppd_depan')
  let cp_sppd_depan = workbook.addWorksheet('cp')

  cp_sppd_depan.model = Object.assign(sppd_depan.model, {
    mergeCells: sppd_depan.model.merges
  })
  cp_sppd_depan.name = `sppd_depan ${formated_date}`
  cp_sppd_depan.getCell('D22').value = formated_date
  cp_sppd_depan.getCell('D22').numFmt = '[$-id-ID]dd mmmm yyyy@'

  // cp sppd_belakang
  const sppd_belakang = workbook.getWorksheet('sppd_belakang')
  let cp_sppd_belakang = workbook.addWorksheet('cp')

  cp_sppd_belakang.model = Object.assign(sppd_belakang.model, {
    mergeCells: sppd_belakang.model.merges
  });
  cp_sppd_belakang.name = `sppd_belakang ${formated_date}`;
  cp_sppd_belakang.getCell('F7').value = formated_date
  cp_sppd_belakang.getCell('F7').numFmt = '[$-id-ID]dd mmmm yyyy;@'
}
</script>

<template>
  
  <form>
    <div class="grid gap-6 mb-6 md:grid-cols-2">
      <div>
        <label for="nama" class="block mb-2 text-sm font-medium text-gray-900 dark:text-white">Nama</label>
        <input v-model="nama" type="text" id="nama" name="nama" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2.5 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-blue-500 dark:focus:border-blue-500" placeholder="Angger" required>
      </div>
      <div>
        <label for="nip" class="block mb-2 text-sm font-medium text-gray-900 dark:text-white">NIP</label>
        <input v-model="nip" type="text" id="nip" name="nama" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2.5 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-blue-500 dark:focus:border-blue-500" placeholder="1994030500000000" required>
      </div>
      <div>
        <label for="golongan" class="block mb-2 text-sm font-medium text-gray-900 dark:text-white">Pangkat Golongan</label>
        <input v-model="golongan" type="text" id="golongan" name="golongan" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2.5 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-blue-500 dark:focus:border-blue-500" placeholder="IVd" required>
      </div>
      <div>
        <label for="jabatan" class="block mb-2 text-sm font-medium text-gray-900 dark:text-white">Jenis</label>
        <input v-model="jabatan" type="text" id="jabatan" name="jabatan" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2.5 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-blue-500 dark:focus:border-blue-500" placeholder="Epidemolog" required>
      </div>
    </div>
    <div class="grid gap-6 mb-6 md:grid-cols-2">
      <div class="mb-6">
        <label for="tanggal" class="block mb-2 text-sm font-medium text-gray-900 dark:text-white">Tanggal</label>
        <VueDatePicker id="tanggal" name="tanggal" v-model="tanggal" text-input multi-dates :enable-time-picker="false" auto-apply :close-on-auto-apply="false" placeholder="mm/dd/yyyyy;" required/>
      </div> 
      <div class="mb-6">
        <label for="alamat" class="block mb-2 text-sm font-medium text-gray-900 dark:text-white">Alamat</label>
        <input v-model="alamat" type="text" id="alamat" name="alamat" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2.5 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-blue-500 dark:focus:border-blue-500" placeholder="Desa Branjang" required>
      </div>
    </div>
      <div class="mb-6">
        <label for="tujuan" class="block mb-2 text-sm font-medium text-gray-900 dark:text-white">Tujuan</label>
        <input v-model="tujuan" type="text" id="tujuan" name="tujuan" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2.5 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-blue-500 dark:focus:border-blue-500" placeholder="Pemantauan Jentik Nyamuk" required>
      </div>
    <button @click="save" type="button" class="text-white bg-blue-700 hover:bg-blue-800 focus:ring-4 focus:outline-none focus:ring-blue-300 font-medium rounded-lg text-sm w-full sm:w-auto px-5 py-2.5 text-center dark:bg-blue-600 dark:hover:bg-blue-700 dark:focus:ring-blue-800">Buat</button>
  </form>

</template>

