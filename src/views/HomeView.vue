<script setup lang="ts">
import { ref, computed, onBeforeMount, onBeforeUnmount, nextTick, watchEffect } from 'vue'
import ExcelJS from 'exceljs'

const prop = defineProps({
  url: {
    type: String,
    required: true,
    default: 'http://localhost:1017/test/sse'
  },
  total: {
    type: Number,
    required: true,
    default: 5
  },
  filename: {
    type: String,
    required: true
  }
})

const states = {
  READY: 'READY',
  DOWNLOADING: 'DOWNLOADING',
  CANCEL: 'CANCEL',
  FINISH: 'FINISH'
}

let enable = ref(true)
/**
 * @type {import('vue').Ref<EventSource>}
 */
let evtSource = ref()
const workbook = ref()
let downloadUrl = ref('')
let number = ref(0)
let state = ref(states.READY)
/**
 * @type {import('vue').Ref<HTMLAnchorElement>}
 */
const downloadRef = ref()

let percent = computed(() => {
  return number.value >= prop.total ? 100 : Math.floor((number.value / prop.total) * 100)
})

watchEffect(() => {
  if (percent.value === 100) {
    downloadExcel()
    closeEventSource()
    state.value = states.FINISH
  }
})

function startDownload(refresh = false) {
  if (refresh) {
    _reset()
  }
  state.value = states.DOWNLOADING
  createEventSource()
  createWorkBook()
}

function createWorkBook() {
  workbook.value = new ExcelJS.Workbook()
}

function createEventSource() {
  const eventSource = new EventSource(prop.url, {
    withCredentials: false
  })
  eventSource.onmessage = onEvent
  eventSource.onerror = () => {
    closeEventSource()
  }
  evtSource.value = eventSource
}

function closeEventSource(manualClose = false) {
  evtSource.value.close()
  state.value = manualClose ? states.CANCEL : states.FINISH
}

/**
 * @type {(event: MessageEvent) => void}
 */
function onEvent(event) {
  const maxSize = 10000
  let sheet = createSheet(`0 - ${maxSize}`)
  if (number.value !== 0 && number.value % maxSize === 0) {
    const current = Math.floor(number.value / maxSize)
    sheet = createSheet(`${current * maxSize} - ${(current + 1) * maxSize}`)
  }
  number.value++
  const data = event.data
  sheet.addRow(JSON.parse(data))
}

function downloadExcel() {
  workbook.value.xlsx
    .writeBuffer()
    .then((buff) => {
      const blob = new Blob([buff], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      })
      downloadUrl.value = URL.createObjectURL(blob)
      nextTick(() => {
        downloadRef.value?.click()
      })
    })
    .catch((error) => {
      console.log(error)
    })
}

/**
 * @type {(sheetName: string) => import('exceljs').Worksheet}
 */
function createSheet(sheetName) {
  const sheet = workbook.value.getWorksheet(sheetName) ?? workbook.value.addWorksheet(sheetName)
  sheet.columns = [
    { header: '姓名', key: 'name' },
    { header: '年龄', key: 'age' },
    { header: '性别', key: 'sex' },
    { header: '地址', key: 'address' }
  ]
  return sheet
}

function _reset() {
  number.value = 0
  workbook.value = new ExcelJS.Workbook()
  downloadUrl.value = ''
  evtSource.value = null
}

function _checkoutAvailable() {
  return fetch(prop.url)
    .then(() => (enable.value = true))
    .catch(() => (enable.value = false))
}

onBeforeMount(async () => {
  await _checkoutAvailable()
})

onBeforeUnmount(() => {
  state.value = states.READY
  _reset()
})
</script>

<template>
  <div class="h-screen w-screen flex">
    <div class="m-auto w-[500px] rounded-lg bg-white ring-1 ring-slate-900/5 shadow-lg space-y-3">
      <div v-if="enable" class="flex p-4 flex-col w-full">
        <!-- 进度条 -->
        <div class="w-full h-3 bg-slate-200 rounded-lg">
          <div
            style="transition: width 0.3s ease-in-out"
            :style="{ width: `${percent}%` }"
            class="h-full bg-pink-500 rounded-lg"
          ></div>
        </div>
        <a
          ref="downloadRef"
          style="visibility: hidden"
          :href="downloadUrl"
          :download="prop.filename"
        ></a>
        <template v-if="state === states.READY">
          <div style="padding: 10px 0">
            <span>等待下载开始</span>
          </div>
          <div class="self-end">
            <button class="rounded-lg p-2 bg-pink-500 text-cyan-50" @click="startDownload(false)">
              开始下载
            </button>
          </div>
        </template>
        <template v-else-if="state === states.DOWNLOADING">
          <div style="padding: 10px 0">
            <span>正在下载: {{ percent }}%</span>
          </div>
          <div class="self-end">
            <button class="rounded-lg p-2 bg-pink-500 text-cyan-50" @click="closeEventSource(true)">
              取消下载
            </button>
          </div>
        </template>
        <template v-else-if="state === states.CANCEL">
          <div style="padding: 10px 0">
            <span>下载已取消: {{ percent }}%</span>
          </div>
          <div class="self-end">
            <button
              class="rounded-lg m-2 p-2 bg-pink-500 text-cyan-50 px-4"
              @click="downloadExcel()"
            >
              下载已获取的数据
            </button>
            <button class="rounded-lg p-2 bg-pink-500 text-cyan-50" @click="startDownload(true)">
              重新下载
            </button>
          </div>
        </template>
        <template v-else-if="state === states.FINISH">
          <div style="padding: 10px 0">
            <span>下载完成</span>
          </div>
          <div class="self-end">
            <button class="rounded-lg p-2 bg-pink-500 text-cyan-50" @click="startDownload(true)">
              重新下载
            </button>
          </div>
        </template>
      </div>
      <div class="flex p-4 flex-col items-center" v-else>
        <span> 网络故障或请求被阻止, 检查网络或者url是否可用 </span>
        <a :href="prop.url">{{ prop.url }}</a>
        <div>
          <button class="rounded-lg p-2 bg-pink-500 text-cyan-50" @click="_checkoutAvailable">
            重新检测
          </button>
        </div>
      </div>
    </div>
  </div>
</template>
