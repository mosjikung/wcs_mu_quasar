<template>
    <div class="q-pa-md" style="max-width: 100%">
    <div class="q-gutter-md">
             <q-tabs v-model="tab" dense class="bg-grey-3" align="justify" narrow-indicator>
                            <q-tab name="main" label="MAIN" />
                            <q-tab name="fabric" label="Fabric" />
                             <q-tab name="fabric" label="QTY" />
                              <q-tab name="fabric" label="AVG" />
                               <q-tab name="fabric" label="Splice Loss" />
                        </q-tabs>
                        
                         <q-tab-panels v-model="tab" animated>
                             <q-tab-panel name="main">
                                <div class="q-pt-md" style="margin-top: 0px;padding-top: 0px;">
                                    <div class="row">
                                   
                                    
                                        <div class="col-4">
                                             <q-input outlined v-model="text" :dense="dense" placeholder="GM no" lable="GM NO" />

                                        </div>
                                        <div class="col-4">

                                        </div>
                                         <div class="col-4">
                                             <q-input outlined v-model="text" :dense="dense" placeholder="Date" lable="GM NO" />

                                        </div>
                                        
                                   
                                    </div>
                                </div>
                            </q-tab-panel>
                            <q-tab-panel name="fabric">
                                <div class="q-pt-md" style="margin-top: 0px;padding-top: 0px;">
                                    <q-table :rows="rows_barcode" row-key="barcode" :pagination="pagination" :rows-per-page-options="[3]" style="min-height: 250px;" />
                                </div>
                            </q-tab-panel>
                                
                         </q-tab-panels>

    </div>
    </div>
</template>

<script>
import {ref} from 'vue'
export default {
setup(){
    return{
       tab: ref('main'),
    }
}
}
</script>

<style>

</style>