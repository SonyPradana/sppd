<script async setup>
import { ref } from 'vue'
import VueDatePicker from '@vuepic/vue-datepicker';
import '@vuepic/vue-datepicker/dist/main.css'
import { save } from '../../sppd/ukm'

const nama = ref([])
const nip = ref([])
const golongan = ref([])
const jabatan = ref([])
const tanggal = ref()
const alamat = ref([])
const tujuan = ref([])
const multy = ref(false)

function isMulty(e) {
  multy.value = ! multy.value
}

async function download() {
  await save(tanggal, {
    nama : nama.value,
    nip: nip.value,
    golongan: golongan.value,
    jabatan: jabatan.value,
    tujuan: tujuan.value,
    alamat: alamat.value
  })
}
</script>

<template>
  
  <form>
    <div class="grid gap-6 mb-6 md:grid-cols-2">
      <div>
        <label for="nama" class="block mb-2 text-sm font-medium text-gray-900 dark:text-white">Nama</label>
        <input v-model="nama[1]" type="text" id="nama" name="nama" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2.5 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-blue-500 dark:focus:border-blue-500" placeholder="Angger" required>
      </div>
      <div>
        <label for="nip" class="block mb-2 text-sm font-medium text-gray-900 dark:text-white">NIP</label>
        <input v-model="nip[1]" type="text" id="nip" name="nama" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2.5 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-blue-500 dark:focus:border-blue-500" placeholder="1994030500000000" required>
      </div>
      <div>
        <label for="golongan" class="block mb-2 text-sm font-medium text-gray-900 dark:text-white">Pangkat Golongan</label>
        <input v-model="golongan[1]" type="text" id="golongan" name="golongan" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2.5 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-blue-500 dark:focus:border-blue-500" placeholder="IVd" required>
      </div>
      <div>
        <label for="jabatan" class="block mb-2 text-sm font-medium text-gray-900 dark:text-white">Jenis</label>
        <input v-model="jabatan[1]" type="text" id="jabatan" name="jabatan" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2.5 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-blue-500 dark:focus:border-blue-500" placeholder="Epidemolog" required>
      </div>
    </div>
    <div v-if="multy" class="grid gap-6 mb-6 md:grid-cols-2">
      <div>
        <label for="nama2" class="block mb-2 text-sm font-medium text-gray-900 dark:text-white">Nama (2)</label>
        <input v-model="nama[2]" type="text" id="nama2" name="nama2" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2.5 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-blue-500 dark:focus:border-blue-500" placeholder="Pradana" required>
      </div>
      <div>
        <label for="nip2" class="block mb-2 text-sm font-medium text-gray-900 dark:text-white">NIP (2)</label>
        <input v-model="nip[2]" type="text" id="nip2" name="nama2" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2.5 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-blue-500 dark:focus:border-blue-500" placeholder="1994030500000000" required>
      </div>
      <div>
        <label for="golongan2" class="block mb-2 text-sm font-medium text-gray-900 dark:text-white">Pangkat Golongan (2)</label>
        <input v-model="golongan[2]" type="text" id="golongan2" name="golongan2" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2.5 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-blue-500 dark:focus:border-blue-500" placeholder="IVa" required>
      </div>
      <div>
        <label for="jabatan2" class="block mb-2 text-sm font-medium text-gray-900 dark:text-white">Jenis (2)</label>
        <input v-model="jabatan[2]" type="text" id="jabatan2" name="jabatan2" class="bg-gray-50 border border-gray-300 text-gray-900 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-2.5 dark:bg-gray-700 dark:border-gray-600 dark:placeholder-gray-400 dark:text-white dark:focus:ring-blue-500 dark:focus:border-blue-500" placeholder="Staff" required>
      </div>
    </div>
    
    <div class="flex justify-center">
      <button @click="isMulty" class="dark:bg-gray-700 dark:text-gray-50 bg-gray-50 text-gray-700">{{ multy ? 1 : 2}} Orang</button>
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
    <button @click="download" type="button" class="text-white bg-blue-700 hover:bg-blue-800 focus:ring-4 focus:outline-none focus:ring-blue-300 font-medium rounded-lg text-sm w-full sm:w-auto px-5 py-2.5 text-center dark:bg-blue-600 dark:hover:bg-blue-700 dark:focus:ring-blue-800">Buat</button>
  </form>

</template>

