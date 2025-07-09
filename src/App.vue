<template>
  <template v-if="loading">
    <Screen v-if="screening" />
    <Editor v-else-if="_isPC" />
    <Mobile v-else />
  </template>
  <FullscreenSpin :tip=tip v-else  loading :mask="false" />
</template>

<script lang="ts" setup>
import { onMounted,nextTick,ref } from 'vue'
import { storeToRefs } from 'pinia'
import { useScreenStore, useMainStore, useSnapshotStore, useSlidesStore } from '@/store'
import { LOCALSTORAGE_KEY_DISCARDED_DB } from '@/configs/storage'
import { deleteDiscardedDB } from '@/utils/database'
import { isPC } from '@/utils/common'
import api from '@/services'
import useAIPPT from '@/hooks/useAIPPT'
import type { AIPPTSlide } from '@/types/AIPPT'

import Editor from './views/Editor/index.vue'
import Screen from './views/Screen/index.vue'
import Mobile from './views/Mobile/index.vue'
import FullscreenSpin from '@/components/FullscreenSpin.vue'
import useSlideHandler from '@/hooks/useSlideHandler'

const _isPC = isPC()

const mainStore = useMainStore()
const slidesStore = useSlidesStore()
const snapshotStore = useSnapshotStore()
const { databaseId } = storeToRefs(mainStore)
const { slides } = storeToRefs(slidesStore)
const { screening } = storeToRefs(useScreenStore())
const loading = ref(false)
const { AIPPT } = useAIPPT()
const { resetSlides } = useSlideHandler()
const tip = ref("数据初始化中，请稍等 ...")

if (import.meta.env.MODE !== 'development') {
  window.onbeforeunload = () => false
}

// 获取URL参数中的url属性
const getUrlParam = (name: string): string | null => {
  const urlParams = new URLSearchParams(window.location.search)
  return urlParams.get(name)
}

// 修改后的 handleAIPPTConfig 方法
const handleAIPPTConfig = async (configData: any) => {
  // 设置基础配置
  mainStore.setAIPPTModel(configData.model || 'gpt-4o')
  mainStore.setAIPPTTemplate(configData.template || 'template_1')
  slidesStore.setTitle(configData.title || 'AIPPT Presentation')
  
  // 如果是需求数据，进入setup步骤
  if (configData.type=="aippt_demand" && configData.demand) {
    mainStore.setAIPPTDemand(configData.demand)
    mainStore.setAIPPTStep('setup')
    await mainStore.setAIPPTDialogState(true)
    await nextTick()
  }
  
  // 如果是大纲数据，进入outline步骤
  else if (configData.type=="aippt_outline" && configData.outline) {
    mainStore.setAIPPTDemand(configData.demand || '') // 保存原始需求（如果有）
    mainStore.setAIPPTOutline(configData.outline)
    mainStore.setAIPPTStep('outline')
    await mainStore.setAIPPTDialogState(true)
    await nextTick()
  }
  
  // 如果是AIPPT JSON数据，直接生成PPT
  else if (configData.type=="aippt_slides" && configData.slides) {
    const templateData = await api.getFileData(configData.template || 'template_1')
    const templateSlides = templateData.slides
    const templateTheme = templateData.theme
    
    // 设置模板主题
    slidesStore.setTheme(templateTheme)
    // 处理AIPPT数据生成PPT
    await processAIPPTSlides(configData.slides, templateSlides)
  }
}

// 处理AIPPT slides数据
const processAIPPTSlides = async (aipptSlides: AIPPTSlide[], templateSlides: any[]) => {
  try {
    // 使用AIPPT hook处理数据
    for (const slide of aipptSlides) {
      await AIPPT(templateSlides, [slide])
    }
  } catch (error) {
    console.error('处理AIPPT slides时出错:', error)
  }
}

// 检查是否为AIPPT配置
const isAIPPTConfig = (data: any): boolean => {
  return data.type &&(data.demand || data.outline || data.slides)
}

onMounted(async () => {
  try {
    resetSlides() // 重置幻灯片
    // 获取配置URL参数
    const configUrl = getUrlParam('aippt_config_url')
    
    if (configUrl) {
      // 如果有配置URL，优先处理配置
      const configData = await api.getUrl(configUrl)
      
      if (isAIPPTConfig(configData)) {
        // 处理AIPPT配置
        await handleAIPPTConfig(configData)
      }
    }else{
      const slideUrl = getUrlParam('slides_url') || './mocks/slides.json'
      const initSlides = await api.getUrl(slideUrl)
      slidesStore.setSlides(initSlides.slides)
      slidesStore.setTitle(initSlides.title)
      slidesStore.setTheme(initSlides.theme)
    }
    loading.value = true
  } catch (error) {
    tip.value = '初始化时出错'
    console.error('初始化时出错:', error)
  }
  await deleteDiscardedDB()
  snapshotStore.initSnapshotDatabase()
})

// 应用注销时向 localStorage 中记录下本次 indexedDB 的数据库ID，用于之后清除数据库
window.addEventListener('beforeunload', () => {
  const discardedDB = localStorage.getItem(LOCALSTORAGE_KEY_DISCARDED_DB)
  const discardedDBList: string[] = discardedDB ? JSON.parse(discardedDB) : []

  discardedDBList.push(databaseId.value)

  const newDiscardedDB = JSON.stringify(discardedDBList)
  localStorage.setItem(LOCALSTORAGE_KEY_DISCARDED_DB, newDiscardedDB)
})
</script>

<style lang="scss">
#app {
  height: 100%;
}
</style>