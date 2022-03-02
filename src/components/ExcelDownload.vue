<template>
    <div>
        <button @click.prevent="excelDownFile">다운로드</button>
    </div>
</template>

<script>
import {createXlsx} from '../common/excelDownloader.js';

export default {
    name: 'ExcelDownload',
    setup() {
        // 서버에서 추가정보 리스트 가져오기
        const extraList = ['djlk','dkl'];

        function excelDownFile() {
            console.log('down...')
            createXlsx(extraList);
        }

        return {
            excelDownFile,
        };
    },
};
</script>

<style>

</style>