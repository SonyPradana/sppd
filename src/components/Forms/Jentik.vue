<script async setup>
import * as ExcelJS from 'exceljs'
import excel from '../../assets/sppd.xlsx'
import axios from 'axios'
import { Buffer } from 'buffer'
import { ref } from 'vue'

const nama = ref('')
const nip = ref('')
const golongan = ref('')
const jabatan = ref('')
const tanggal = ref('')
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
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(await getResponseAsBuffer(excel))

  const ws = workbook.getWorksheet('Surat tugas 1 org')
  ws.getCell('E13').value = nama.value
  ws.getCell('E14').value = nip.value
  ws.getCell('E15').value = golongan.value
  ws.getCell('E16').value = jabatan.value
  ws.getCell('E23').value = tanggal.value
  ws.getCell('E24').value = tujuan.value
  ws.getCell('E25').value = alamat.value

  const buffer = await workbook.xlsx.writeBuffer()
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  
  const link = document.createElement('a')
  link.href = window.URL.createObjectURL(blob)
  link.download = `sppd ${nama.value}.xlsx`
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)

  busy.value = false
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
        <input v-model="tanggal" type="text" id="tanggal" name="tanggal" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2.5 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-blue-500 dark:focus:border-blue-500" placeholder="dd-mm-yyyyy" required>
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

