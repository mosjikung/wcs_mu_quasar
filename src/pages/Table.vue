<template>
   <div class="row">
                <div class="col-12">
                <div class="q-pt-md" style="margin-top: 0px;padding-top: 0px;">
                                        <q-table :rows="rowshow"  :columns="columnsshow"  table-header-class="bg-blue-8" :rows-per-page-options="[10]" row-key="no" style="min-height: 200px; blackground-color:black border: 0px;">
                                            <template v-slot:body="props">
                                                <q-tr  :props="props">
                                                  
                                                        <!-- <q-btn label="Cnacel Barcode" color="negative" @click="CANCEL_BARCODE" size="12px" padding="4px"></q-btn> -->
                                                        <!--   -->
                                                      
                                               
                                                    
                                                     <q-td key="no" :props="props">
                                                      <q-input input-class="text-center" v-model="props.row.no"  :dense="dense" disable />

                                                    </q-td>
                                                  
                                                    <q-td key="so" :props="props">
                                                        <q-input input-class="text-center" v-model="props.row.so" :dense="dense" disable />
                                                    </q-td>
                                                     <q-td key="year" :props="props">
                                                     <q-input input-class="text-center" v-model="props.row.year"  :dense="dense" disable />
                                                    </q-td>
                                                     <q-td key="style" :props="props">
                                                        <q-input input-class="text-center" v-model="props.row.style"  :dense="dense" disable />
                                                    </q-td>
                                                    <q-td key="customer" :props="props"> <!--key = name.columne -->
                                                       <q-input input-class="text-center"    v-model="props.row.customer"  :dense="dense" disable />
                                                    </q-td>
                                                    <q-td key="customer" :props="props"> <!--key = name.columne -->
                                                       <q-btn  color="amber-13" icon="my_location" @click="$router.push('pages/new')"/>
                                                    </q-td>
                                                     <q-td key="customer" :props="props"> <!--key = name.columne -->
                                                       <q-btn  color="red-12" icon="delete" @click="deleteSelected(props.pageIndex)"/>
                                                    </q-td>
                                                   
                                                   
                                                  

                                                    
                                          
                                                </q-tr>
                                            </template>
                                        </q-table>
                                       
                                       
                                    </div>
                </div>

                

        </div>
</template>

<script>
import {ref} from 'vue'
import {date} from 'quasar'
import { colors } from 'quasar'
export default {
    
    setup(){
       const { getPaletteColor } = colors
    
        return{
            date: ref('2021-10-21')
        }
    },
        data(){
            return{
                search:"",
                columnsshow:[
                    {
                        name:"no",
                        label:"NO.",
                        align:"center",
                        field:"no"
                    },
                    {
                        name:"so",
                        label:"SO",
                        align:"center",
                        field:"no"
                    },
                    {
                        name:"year",
                        label:"Year",
                        align:"center",
                        field:"year"
                    },
                    {
                        name:"style",
                        label:"Style",
                        align:"center",
                        field:"style"
                    },
                    {
                        name:"customer",
                        label:"Customer",
                        align:"center",
                        field:"customer"
                    },
                     {
                        name:"detail",
                        label:"Detail",
                        align:"center",
                        field:"detail"
                    },
                    {
                        name:"delete",
                        label:"Delete",
                        align:"center",
                        field:"delete"
                    },

                    

                ],
                rowshow:[
                    {
                        no:'1',
                        so:'01212536',
                        year:'21',
                        style:'w5396',
                        customer:'BCS'
                    }
                ]

            }
            
        },
        methods: {
            checksearch(){
                console.log(this.search)
            },
            addRow(){
                this.rowshow.push([
                    {
                         no:'',
                        so:'',
                        year:'',
                        style:'',
                        customer:''
                    }
                ])
            },
            removeRow(){
                this.rowshow.pop()
            },
            deleteSelected(index){
           console.log(index)
           this.rowshow[index].no='';
           this.rowshow[index].so='';
           this.rowshow[index].year='';
           this.rowshow[index].style='';
           this.rowshow[index].customer='';
      
        
    },
             
        }
        
}
</script>

