<template>
  <q-btn @click="check()" label="check"></q-btn>
</template>

<script>
export default {
    data(){
      return{
          information:[
            {
              name:'weerayooth',
              sname:'kewruksa',
              age:'31'
            },
            {
              name:'testuser',
              sname:'testname',
              age:'25'
            }
          ]
      }
    },
    methods: {
            check(){
              this.$emit('sentdata',this.information)
            }
    },
}
</script>

<style>

</style>