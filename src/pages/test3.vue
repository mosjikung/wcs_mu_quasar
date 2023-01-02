<template>
  
</template>

<script>
export default {
//Daily

    async  exportexcel(){
      
      
         if(count == 0){
         
       worksheet.getCell("A"+a).value = "Date";
      worksheet.mergeCells("A"+a +":" + "A"+b);
      worksheet.getCell("A"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "A" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].mu_date;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
      var ty = this.length
    
       worksheet.getCell("A"+ty).value = "Total Type";
      worksheet.mergeCells("A"+ty +":" + "I"+ty);
      worksheet.getCell("A"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
        this.result_ttlmarker = 0
      for(var re_ttlmarker = 0; re_ttlmarker<resp.data.data.length; re_ttlmarker++ ){
        this.result_ttlmarker =  Number(this.result_ttlmarker) + Number(this.rowexport[re_ttlmarker].ttl_marker)
        
        
      }
      this.keep_ttlmarker.push({
          result_ttlmarker:this.result_ttlmarker,
        })
        worksheet.getCell("J"+ty).value = this.result_ttlmarker;
      worksheet.getCell("J"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_plug12 = 0
      for(var re_plug1_2 = 0; re_plug1_2<resp.data.data.length; re_plug1_2++ ){
        this.result_plug12 =  Number(this.result_plug12) + Number(this.rowexport[re_plug1_2].plug12)
        
      }
     
     this.result_plug12 = this.result_plug12/resp.data.data.length 
     this.keep_plug12_inch.push({
          result_plug12_inch:this.result_plug12,
        })

        worksheet.getCell("K"+ty).value = parseFloat(this.result_plug12).toFixed(2);
      worksheet.getCell("K"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 

      
      this.result_perplug12 = 0
      for(var re_perplug12 = 0; re_perplug12<resp.data.data.length; re_perplug12++ ){
        this.result_perplug12 =  Number(this.result_perplug12) + Number(this.rowexport[re_perplug12].widthloss)
        
      }
      this.keep_perplug12.push({
          result_perplug12:this.result_perplug12,
        })

     
     this.result_perplug12 = this.result_perplug12/resp.data.data.length 

        worksheet.getCell("L"+ty).value = this.result_perplug12;
      worksheet.getCell("L"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
      
      for(var replug12_inch = 0; replug12_inch<resp.data.data.length; replug12_inch++){
      this.result_plug12_inch = (this.rowexport[replug12_inch].endless_length_yd*36)/(this.rowexport[replug12_inch].ttl_marker/this.rowexport[replug12_inch].length_ydb)
    
      this.end_loss_plug12.push(this.result_plug12_inch)
        
      }
        this.result_plug12_avg = 0;
        for(var replug12_avg = 0; replug12_avg<this.end_loss_plug12.length; replug12_avg++){
        this.result_plug12_avg= Number(this.result_plug12_avg) + Number(this.end_loss_plug12[replug12_avg])
      
      }  
     
        this.all_result_avg = ((this.result_plug12_avg * this.result_ttlmarker) / this.result_ttlmarker)
      
      
        
      worksheet.getCell("M"+ty).value = parseFloat(this.all_result_avg).toFixed(2);
      worksheet.getCell("M"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_length_yd = 0
      for(var re_length_yd = 0; re_length_yd<resp.data.data.length; re_length_yd++ ){
        this.result_length_yd =  Number(this.result_length_yd) + Number(this.rowexport[re_length_yd].endless_length_yd)
        
      }
     this.sum_length_yd = this.result_length_yd
     this.result_length_yd = this.result_length_yd/resp.data.data.length 

        worksheet.getCell("N"+ty).value = parseFloat(this.result_length_yd).toFixed(2);
      worksheet.getCell("N"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


       this.result_avg_per = 0;
      this.result_avg_per = this.result_length_yd / this.result_ttlmarker
      
      this.keep_avg_end.push({
       result_avg_per: this.result_avg_per
      })
      
      worksheet.getCell("O"+ty).value = parseFloat(this.result_avg_per).toFixed(2);
      worksheet.getCell("O"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_splice = 0
      for(var re_splice = 0; re_splice<resp.data.data.length; re_splice++ ){
        this.result_splice =  Number(this.result_splice) + Number(this.rowexport[re_splice].spliceloss)
        
      }
     this.sum_result_splice = this.result_splice
     this.result_splice = this.result_splice/resp.data.data.length
     this.keep_splice_Loss.push({
          result_splice:this.result_splice
     })


        worksheet.getCell("P"+ty).value = parseFloat(this.result_splice).toFixed(2);
      worksheet.getCell("P"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_splicelossper = 0;
      this.result_splicelossper = this.result_splice / this.result_ttlmarker
      worksheet.getCell("Q"+ty).value = parseFloat(this.result_splicelossper).toFixed(2);
      worksheet.getCell("Q"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


       this.result_cut_out_loss = 0
      for(var re_cut_out = 0; re_cut_out<resp.data.data.length; re_cut_out++ ){
        this.result_cut_out_loss =  Number(this.result_cut_out_loss) + Number(this.rowexport[re_cut_out].totalcutout)
        
      }
     this.sum_cut_loss_out = this.result_cut_out_loss
     this.result_cut_out_loss = this.result_cut_out_loss/resp.data.data.length 
     this.keep_cut_out_loss.push({
       result_cut_out_loss: this.result_cut_out_loss
     })

        worksheet.getCell("R"+ty).value = parseFloat(this.result_cut_out_loss).toFixed(2);
      worksheet.getCell("R"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
      this.result_cut_out_loss_per = 0;
      this.result_cut_out_loss_per = this.result_cut_out_loss / this.result_ttlmarker
      worksheet.getCell("S"+ty).value = parseFloat(this.result_cut_out_loss_per).toFixed(2);
      worksheet.getCell("S"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_remnants = 0
      for(var re_remnants = 0; re_remnants<resp.data.data.length; re_remnants++ ){
        this.result_remnants =  Number(this.result_remnants) + Number(this.rowexport[re_remnants].rement)
        
      }
     this.sum_remnants = this.result_remnants
     this.result_remnants = this.result_remnants/resp.data.data.length
     this.keep_remnants.push({
       result_remnants:this.result_remnants
     })

        worksheet.getCell("T"+ty).value = parseFloat(this.result_remnants).toFixed(2);
      worksheet.getCell("T"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_remnants_per = 0;
      this.result_remnants_per = this.result_remnants / this.result_ttlmarker
      worksheet.getCell("U"+ty).value = parseFloat(this.result_remnants_per).toFixed(2);
      worksheet.getCell("U"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


      this.result_total_length = 0
      for(var re_total_length = 0; re_total_length<resp.data.data.length; re_total_length++ ){
      this.result_total_length =  Number(this.result_total_length) + Number(this.rowexport[re_total_length].totallenthloss)
        
      }
     this.sum_total_length = this.result_total_length
     this.result_total_length = this.result_total_length/resp.data.data.length
     this.keep_total_loss.push({
       result_total_length:this.result_total_length
     }) 

        worksheet.getCell("V"+ty).value = parseFloat(this.result_total_length).toFixed(2);
      worksheet.getCell("V"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       this.result_total_length_per = 0;
      this.result_total_length_per = this.result_total_length / this.result_ttlmarker
      worksheet.getCell("W"+ty).value = parseFloat(this.result_total_length_per).toFixed(2);
      worksheet.getCell("W"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };








       //B---------------------------------------------------------------------
       worksheet.getCell("B"+a).value = "FL";
      worksheet.mergeCells("B"+a +":" + "B"+b);
      worksheet.getCell("B"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("B"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "B" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_table;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

        //C------------------------------------------------------------------------
        worksheet.getCell("C"+a).value = "S/O no.";
      worksheet.mergeCells("C"+a +":" + "C"+b);
      worksheet.getCell("C"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("C"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "C" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_no_short;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

        //D------------------------------------------------------------------------
        worksheet.getCell("D"+a).value = "Customer";
      worksheet.mergeCells("D"+a +":" + "D"+b);
      worksheet.getCell("D"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("D"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "D" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].customer;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //E-----------------------------------------------------------
      worksheet.getCell("E"+a).value = "GM08 no.";
      worksheet.mergeCells("E"+a +":" + "E"+b);
      worksheet.getCell("E"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("E"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "E" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_no;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //F-------------------------------------------------------------
       if(this.multiple[x] == 1){
         this.name_f_type = "Woven"
       }
       if(this.multiple[x] == 2){
         this.name_f_type = "Light-Middle"
       }
      
      if(this.multiple[x] == 3){
         this.name_f_type = "Fleece"
       }
      if(this.multiple[x] == 4){
         this.name_f_type = "Special"
       }
      worksheet.getCell("F"+a).value = this.name_f_type;
     worksheet.mergeCells("F"+a +":" + "F"+b);
      worksheet.getCell("F"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("F"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "F" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].fabric_type;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
     
      //G----------------------------------------------------------------------

       worksheet.getCell("G"+a).value = "Observe";
      worksheet.getCell("G"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("G"+b).value = "Width";
      worksheet.getCell("G"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "G" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].obs_width;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //H-------------------------------------------------------
       worksheet.getCell("H"+a).value = "Width";
      worksheet.getCell("H"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("H"+b).value = "(inch)";
      worksheet.getCell("H"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "H" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].width_inch;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
      //I-----------------------------------------------------------
         worksheet.getCell("I"+a).value = "Length";
      worksheet.getCell("I"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("I"+b).value = "(yd)";
      worksheet.getCell("I"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 

      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "I" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].length_ydb;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


       //J-----------------------------------------------------------
         worksheet.getCell("J"+a).value = "Ttl Marker";
      worksheet.getCell("J"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("J"+b).value = "Length(yd)";
      worksheet.getCell("J"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "J" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].ttl_marker;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

//K-------------------------------------------------------------------------------
         worksheet.getCell("K"+a).value = "Plug1+2";
      worksheet.getCell("K"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("K"+b).value = "(inch)";
      worksheet.getCell("K"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
         this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "K" + ab
      
            var z = ab-a-2;
      worksheet.getCell(y).value = parseFloat(check_null(this.rowexport[z].plug12)).toFixed(2);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

//L------------------------------------------------------------------------------------
         worksheet.getCell("L"+a).value = "%";
      worksheet.getCell("L"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("L"+b).value = "1.50%";
      worksheet.getCell("L"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "L" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].widthloss;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
       
  //M--------------------------------------------------------------
        worksheet.getCell("M"+a).value = "Plug1+2";
      worksheet.getCell("M"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("M"+b).value = "(inch)";
      worksheet.getCell("M"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "M" + ab
      
            var z = ab-a-2;

      this.result_plug12_inch = parseFloat((this.rowexport[z].endless_length_yd*36)/(this.rowexport[z].ttl_marker/this.rowexport[z].length_ydb)).toFixed(2)
       worksheet.getCell(y).value = this.result_plug12_inch;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

    //N----------------------------------------------------------
       worksheet.getCell("N"+a).value = "Length";
      worksheet.getCell("N"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("N"+b).value = "(yd)";
      worksheet.getCell("N"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "N" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].endless_length_yd;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //O------------------------------------------------------------------------
        worksheet.getCell("O"+a).value = "%";
      worksheet.getCell("O"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("O"+b).value = "0.23%";
      worksheet.getCell("O"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "O" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].avg_end+"%";
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
      //P--------------------------------------------------------------------------
         worksheet.getCell("P"+a).value = "Length";
      worksheet.getCell("P"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("P"+b).value = "(yd)";
      worksheet.getCell("P"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "P" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].spliceloss;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //Q------------------------------------------------------------------------------
         worksheet.getCell("Q"+a).value = "%";
      worksheet.getCell("Q"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("Q"+b).value = "0.20%";
      worksheet.getCell("Q"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "Q" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].splicelossper+"%";
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //R------------------------------------------------------------------------------
         worksheet.getCell("R"+a).value = "Length";
      worksheet.getCell("R"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("R"+b).value = "(yd)";
      worksheet.getCell("R"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
          this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "R" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].totalcutout;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


      //S------------------------------------------------------------------------------
         worksheet.getCell("S"+a).value = "%";
      worksheet.getCell("S"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("S"+b).value = "0.00%";
      worksheet.getCell("S"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
        this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "S" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = parseFloat(this.rowexport[z].totalcutoutper).toFixed(2)+"%";
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //T-------------------------------------------------------------------------------

         worksheet.getCell("T"+a).value = "Length";
      worksheet.getCell("T"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("T"+b).value = "(yd)";
      worksheet.getCell("T"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "T" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].rement;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //U-------------------------------------------------------------------------------

         worksheet.getCell("U"+a).value = "%";
      worksheet.getCell("U"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("U"+b).value = "0.14%";
      worksheet.getCell("U"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "U" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].rement_per+"%";
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //V-------------------------------------------------------------------------------

         worksheet.getCell("V"+a).value = "Length";
      worksheet.getCell("V"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("V"+b).value = "(yd)";
      worksheet.getCell("V"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "V" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].totallenthloss;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


       //W-------------------------------------------------------------------------------

         worksheet.getCell("W"+a).value = "%of";
      worksheet.getCell("W"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("W"+b).value = "0.49%";
      worksheet.getCell("W"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "W" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].totallenthlossper+"%";
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
      
        
         } else if(count == 1){
             

             a = this.length+2
             b = this.length+3
     
      worksheet.getCell("A"+a).value = "Date";
      worksheet.mergeCells("A"+a +":" + "A"+b);
      worksheet.getCell("A"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "A" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].mu_date;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
       var ty = this.length+x
    
       worksheet.getCell("A"+ty).value = "Total Type";
      worksheet.mergeCells("A"+ty +":" + "I"+ty);
      worksheet.getCell("A"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
        this.result_ttlmarker = 0
      for(var re_ttlmarker = 0; re_ttlmarker<resp.data.data.length; re_ttlmarker++ ){
        this.result_ttlmarker =  Number(this.result_ttlmarker) + Number(this.rowexport[re_ttlmarker].ttl_marker)
      
      }
        this.keep_ttlmarker.push({
          result_ttlmarker:this.result_ttlmarker,
        })
        worksheet.getCell("J"+ty).value = this.result_ttlmarker;
      worksheet.getCell("J"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };  
      
      this.result_plug12 = 0
      for(var re_plug1_2 = 0; re_plug1_2<resp.data.data.length; re_plug1_2++ ){
        this.result_plug12 =  Number(this.result_plug12) + Number(this.rowexport[re_plug1_2].plug12)
        
      }
     
     this.result_plug12 = this.result_plug12/resp.data.data.length 
     this.keep_plug12_inch.push({
          result_plug12_inch:this.result_plug12,
        })


         worksheet.getCell("K"+ty).value = parseFloat(this.result_plug12).toFixed(2);
      worksheet.getCell("K"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 

       this.result_perplug12 = 0
      for(var re_perplug12 = 0; re_perplug12<resp.data.data.length; re_perplug12++ ){
        this.result_perplug12 =  Number(this.result_perplug12) + Number(this.rowexport[re_perplug12].widthloss)
        
      }
   this.keep_perplug12.push({
          result_perplug12:this.result_perplug12,
        })
     this.result_perplug12_l = this.result_perplug12/resp.data.data.length 
     

        worksheet.getCell("L"+ty).value = this.result_perplug12;
      worksheet.getCell("L"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       for(var replug12_inch = 0; replug12_inch<resp.data.data.length; replug12_inch++){
      this.result_plug12_inch = (this.rowexport[replug12_inch].endless_length_yd*36)/(this.rowexport[replug12_inch].ttl_marker/this.rowexport[replug12_inch].length_ydb)
    
      this.end_loss_plug12.push(this.result_plug12_inch)
        
      }
        this.result_plug12_avg = 0;
        for(var replug12_avg = 0; replug12_avg<this.end_loss_plug12.length; replug12_avg++){
        this.result_plug12_avg= Number(this.result_plug12_avg) + Number(this.end_loss_plug12[replug12_avg])
      
      }  
     
        this.all_result_avg = ((this.result_plug12_avg * this.result_ttlmarker) / this.result_ttlmarker)
        this.keep_avg_loss_plug12.push({
          all_result_avg:this.all_result_avg
        })    

        
      worksheet.getCell("M"+ty).value = parseFloat(this.all_result_avg).toFixed(2);
      worksheet.getCell("M"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_length_yd = 0
      for(var re_length_yd = 0; re_length_yd<resp.data.data.length; re_length_yd++ ){
        this.result_length_yd =  Number(this.result_length_yd) + Number(this.rowexport[re_length_yd].endless_length_yd)
        
      }
     
     this.result_length_yd = this.result_length_yd/resp.data.data.length 

        worksheet.getCell("N"+ty).value = parseFloat(this.result_length_yd).toFixed(2);
      worksheet.getCell("N"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


      this.result_avg_per = 0;
      this.result_avg_per = this.result_length_yd / this.result_ttlmarker
      
      this.keep_avg_end.push({
       result_avg_per: this.result_avg_per
      })
      
      worksheet.getCell("O"+ty).value = parseFloat(this.result_avg_per).toFixed(2);
      worksheet.getCell("O"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_splice = 0
      for(var re_splice = 0; re_splice<resp.data.data.length; re_splice++ ){
        this.result_splice =  Number(this.result_splice) + Number(this.rowexport[re_splice].spliceloss)
        
      }
     
     this.result_splice = this.result_splice/resp.data.data.length
     this.keep_splice_Loss.push({
          result_splice:this.result_splice
     })

        worksheet.getCell("P"+ty).value = parseFloat(this.result_splice).toFixed(2);
      worksheet.getCell("P"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_splicelossper = 0;
      this.result_splicelossper = this.result_splice / this.result_ttlmarker
      worksheet.getCell("Q"+ty).value = parseFloat(this.result_splicelossper).toFixed(2);
      worksheet.getCell("Q"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


       this.result_cut_out_loss = 0
      for(var re_cut_out = 0; re_cut_out<resp.data.data.length; re_cut_out++ ){
        this.result_cut_out_loss =  Number(this.result_cut_out_loss) + Number(this.rowexport[re_cut_out].totalcutout)
        
      }
     
     this.result_cut_out_loss = this.result_cut_out_loss/resp.data.data.length 
     this.keep_cut_out_loss.push({
       result_cut_out_loss: this.result_cut_out_loss
     })

        worksheet.getCell("R"+ty).value = parseFloat(this.result_cut_out_loss).toFixed(2);
      worksheet.getCell("R"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
      this.result_cut_out_loss_per = 0;
      this.result_cut_out_loss_per = this.result_cut_out_loss / this.result_ttlmarker
      worksheet.getCell("S"+ty).value = parseFloat(this.result_cut_out_loss_per).toFixed(2);
      worksheet.getCell("S"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_remnants = 0
      for(var re_remnants = 0; re_remnants<resp.data.data.length; re_remnants++ ){
        this.result_remnants =  Number(this.result_remnants) + Number(this.rowexport[re_remnants].rement)
        
      }
     
     this.result_remnants = this.result_remnants/resp.data.data.length
     this.keep_remnants.push({
       result_remnants:this.result_remnants
     })

        worksheet.getCell("T"+ty).value = parseFloat(this.result_remnants).toFixed(2);
      worksheet.getCell("T"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_remnants_per = 0;
      this.result_remnants_per = this.result_remnants / this.result_ttlmarker
      worksheet.getCell("U"+ty).value = parseFloat(this.result_remnants_per).toFixed(2);
      worksheet.getCell("U"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


      this.result_total_length = 0
      for(var re_total_length = 0; re_total_length<resp.data.data.length; re_total_length++ ){
      this.result_total_length =  Number(this.result_total_length) + Number(this.rowexport[re_total_length].totallenthloss)
        
      }
     
     this.result_total_length = this.result_total_length/resp.data.data.length
     this.keep_total_loss.push({
       result_total_length:this.result_total_length
     }) 

        worksheet.getCell("V"+ty).value = parseFloat(this.result_total_length).toFixed(2);
      worksheet.getCell("V"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       this.result_total_length_per = 0;
      this.result_total_length_per = this.result_total_length / this.result_ttlmarker
      worksheet.getCell("W"+ty).value = parseFloat(this.result_total_length_per).toFixed(2);
      worksheet.getCell("W"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };




       //B---------------------------------------------------------------------

      worksheet.getCell("B"+a).value = "FL";
      worksheet.mergeCells("B"+a +":" + "B"+b);
      worksheet.getCell("B"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("B"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "B" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_table;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //C---------------------------------------------------------------------

       worksheet.getCell("C"+a).value = "S/O no.";
      worksheet.mergeCells("C"+a +":" + "C"+b);
      worksheet.getCell("C"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("C"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "C" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_no_short;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       
       //D---------------------------------------------------------------------

       worksheet.getCell("D"+a).value = "Customer";
      worksheet.mergeCells("D"+a +":" + "D"+b);
      worksheet.getCell("D"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("D"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "D" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].customer;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //E---------------------------------------------------------------------------
       
       worksheet.getCell("E"+a).value = "GM08 no.";
      worksheet.mergeCells("E"+a +":" + "E"+b);
      worksheet.getCell("E"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("E"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "E" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_no;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
       //F---------------------------------------------------------------
       if(this.multiple[x] == 1){
         this.name_f_type = "Woven"
       }
       if(this.multiple[x] == 2){
         this.name_f_type = "Light-Middle"
       }
      
      if(this.multiple[x] == 3){
         this.name_f_type = "Fleece"
       }
      if(this.multiple[x] == 4){
         this.name_f_type = "Special"
       }
    
      worksheet.getCell("F"+a).value = this.name_f_type;
     worksheet.mergeCells("F"+a +":" + "F"+b);
      worksheet.getCell("F"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("F"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "F" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].fabric_type;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //G----------------------------------------------------------
       
       worksheet.getCell("G"+a).value = "Observe";
      worksheet.getCell("G"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("G"+b).value = "Width";
      worksheet.getCell("G"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "G" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].obs_width;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
  //H----------------------------------------------------------
        worksheet.getCell("H"+a).value = "Width";
      worksheet.getCell("H"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("H"+b).value = "(inch)";
      worksheet.getCell("H"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "H" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].width_inch;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //I--------------------------------------------------------------
         worksheet.getCell("I"+a).value = "Length";
      worksheet.getCell("I"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("I"+b).value = "(yd)";
      worksheet.getCell("I"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
      this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "I" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].length_ydb;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
       //J-----------------------------------------------------------
         worksheet.getCell("J"+a).value = "Ttl Marker";
      worksheet.getCell("J"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("J"+b).value = "Length(yd)";
      worksheet.getCell("J"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "J" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].ttl_marker;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
       }; 
       }

        //K-------------------------------------------------------------------------------
         worksheet.getCell("K"+a).value = "Plug1+2";
      worksheet.getCell("K"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("K"+b).value = "(inch)";
      worksheet.getCell("K"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
         
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "K" + ab
      
            var z = ab-a-2;
      worksheet.getCell(y).value = parseFloat(check_null(this.rowexport[z].plug12)).toFixed(2);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
       }; 
       }

       //L------------------------------------------------------------------
            worksheet.getCell("L"+a).value = "%";
      worksheet.getCell("L"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("L"+b).value = "1.50%";
      worksheet.getCell("L"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

          this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "L" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].widthloss;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
       }; 
       }
        //M--------------------------------------------------------------
        worksheet.getCell("M"+a).value = "Plug1+2";
      worksheet.getCell("M"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("M"+b).value = "(inch)";
      worksheet.getCell("M"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "M" + ab
      
            var z = ab-a-2;

      this.result_plug12_inch = parseFloat((this.rowexport[z].endless_length_yd*36)/(this.rowexport[z].ttl_marker/this.rowexport[z].length_ydb)).toFixed(2)
       worksheet.getCell(y).value = this.result_plug12_inch;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

 //N----------------------------------------------------------
       worksheet.getCell("N"+a).value = "Length";
      worksheet.getCell("N"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("N"+b).value = "(yd)";
      worksheet.getCell("N"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "N" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].endless_length_yd;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


       //O------------------------------------------------------------------------
        worksheet.getCell("O"+a).value = "%";
      worksheet.getCell("O"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("O"+b).value = "0.15%";
      worksheet.getCell("O"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "O" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].avg_end+"%";
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
      //P--------------------------------------------------------------------------
         worksheet.getCell("P"+a).value = "Length";
      worksheet.getCell("P"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("P"+b).value = "(yd)";
      worksheet.getCell("P"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "P" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].spliceloss;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //Q------------------------------------------------------------------------------
         worksheet.getCell("Q"+a).value = "%";
      worksheet.getCell("Q"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("Q"+b).value = "0.20%";
      worksheet.getCell("Q"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "Q" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].splicelossper+"%";
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //R------------------------------------------------------------------------------
         worksheet.getCell("R"+a).value = "Length";
      worksheet.getCell("R"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("R"+b).value = "(yd)";
      worksheet.getCell("R"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
          this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "R" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].totalcutout;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


      //S------------------------------------------------------------------------------
         worksheet.getCell("S"+a).value = "%";
      worksheet.getCell("S"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("S"+b).value = "0.00%";
      worksheet.getCell("S"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
        this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "S" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = parseFloat(this.rowexport[z].totalcutoutper).toFixed(2)+"%";
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //T-------------------------------------------------------------------------------

         worksheet.getCell("T"+a).value = "Length";
      worksheet.getCell("T"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("T"+b).value = "(yd)";
      worksheet.getCell("T"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "T" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].rement;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //U-------------------------------------------------------------------------------

         worksheet.getCell("U"+a).value = "%";
      worksheet.getCell("U"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("U"+b).value = "0.14%";
      worksheet.getCell("U"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "U" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].rement_per+"%";
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //V-------------------------------------------------------------------------------

         worksheet.getCell("V"+a).value = "Length";
      worksheet.getCell("V"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("V"+b).value = "(yd)";
      worksheet.getCell("V"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "V" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].totallenthloss;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


       //W-------------------------------------------------------------------------------

         worksheet.getCell("W"+a).value = "%of";
      worksheet.getCell("W"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("W"+b).value = "0.49%";
      worksheet.getCell("W"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "W" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].totallenthlossper+"%";
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

     

         } else{
          

             a = this.length+2
             b = this.length+3
     
      worksheet.getCell("A"+a).value = "Date";
      worksheet.mergeCells("A"+a +":" + "A"+b);
      worksheet.getCell("A"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "A" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].mu_date;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
      
       var ty = this.length+1
    
       worksheet.getCell("A"+ty).value = "Total Type";
      worksheet.mergeCells("A"+ty +":" + "I"+ty);
      worksheet.getCell("A"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
         this.result_ttlmarker = 0
      for(var re_ttlmarker = 0; re_ttlmarker<resp.data.data.length; re_ttlmarker++ ){
        this.result_ttlmarker =  Number(this.result_ttlmarker) + Number(this.rowexport[re_ttlmarker].ttl_marker)
        
      }
      this.keep_ttlmarker.push({
          result_ttlmarker:this.result_ttlmarker,
        })
        worksheet.getCell("J"+ty).value = this.result_ttlmarker;
      worksheet.getCell("J"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

     this.result_plug12 = 0
      for(var re_plug1_2 = 0; re_plug1_2<resp.data.data.length; re_plug1_2++ ){
        this.result_plug12 =  Number(this.result_plug12) + Number(this.rowexport[re_plug1_2].plug12)
        
      }
     
     this.result_plug12 = this.result_plug12/resp.data.data.length 
     this.keep_plug12_inch.push({
          result_plug12_inch:this.result_plug12,
        })


         worksheet.getCell("K"+ty).value = parseFloat(this.result_plug12).toFixed(2);
      worksheet.getCell("K"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 

       this.result_perplug12 = 0
      for(var re_perplug12 = 0; re_perplug12<resp.data.data.length; re_perplug12++ ){
        this.result_perplug12 =  Number(this.result_perplug12) + Number(this.rowexport[re_perplug12].widthloss)
        
      }
       this.keep_perplug12.push({
          result_perplug12:this.result_perplug12,
        })

     
     this.result_perplug12 = this.result_perplug12/resp.data.data.length 

        worksheet.getCell("L"+ty).value = this.result_perplug12;
      worksheet.getCell("L"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       for(var replug12_inch = 0; replug12_inch<resp.data.data.length; replug12_inch++){
      this.result_plug12_inch = (this.rowexport[replug12_inch].endless_length_yd*36)/(this.rowexport[replug12_inch].ttl_marker/this.rowexport[replug12_inch].length_ydb)
    
      this.end_loss_plug12.push(this.result_plug12_inch)
        
      }
        this.result_plug12_avg = 0;
        for(var replug12_avg = 0; replug12_avg<this.end_loss_plug12.length; replug12_avg++){
        this.result_plug12_avg= Number(this.result_plug12_avg) + Number(this.end_loss_plug12[replug12_avg])
      
      }  
     
        this.all_result_avg = ((this.result_plug12_avg * this.result_ttlmarker) / this.result_ttlmarker)
      
      
        
      worksheet.getCell("M"+ty).value = parseFloat(this.all_result_avg).toFixed(2);
      worksheet.getCell("M"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_length_yd = 0
      for(var re_length_yd = 0; re_length_yd<resp.data.data.length; re_length_yd++ ){
        this.result_length_yd =  Number(this.result_length_yd) + Number(this.rowexport[re_length_yd].endless_length_yd)
        
      }
     
     this.result_length_yd = this.result_length_yd/resp.data.data.length 

        worksheet.getCell("N"+ty).value = parseFloat(this.result_length_yd).toFixed(2);
      worksheet.getCell("N"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


      
       this.result_avg_per = 0;
      this.result_avg_per = this.result_length_yd / this.result_ttlmarker
      
      this.keep_avg_end.push({
       result_avg_per: this.result_avg_per
      })
      
      worksheet.getCell("O"+ty).value = parseFloat(this.result_avg_per).toFixed(2);
      worksheet.getCell("O"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_splice = 0
      for(var re_splice = 0; re_splice<resp.data.data.length; re_splice++ ){
        this.result_splice =  Number(this.result_splice) + Number(this.rowexport[re_splice].spliceloss)
        
      }
     
     this.result_splice = this.result_splice/resp.data.data.length
     this.keep_splice_Loss.push({
          result_splice:this.result_splice
     })


        worksheet.getCell("P"+ty).value = parseFloat(this.result_splice).toFixed(2);
      worksheet.getCell("P"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_splicelossper = 0;
      this.result_splicelossper = this.result_splice / this.result_ttlmarker
      worksheet.getCell("Q"+ty).value = parseFloat(this.result_splicelossper).toFixed(2);
      worksheet.getCell("Q"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


       this.result_cut_out_loss = 0
      for(var re_cut_out = 0; re_cut_out<resp.data.data.length; re_cut_out++ ){
        this.result_cut_out_loss =  Number(this.result_cut_out_loss) + Number(this.rowexport[re_cut_out].totalcutout)
        
      }
     
     this.result_cut_out_loss = this.result_cut_out_loss/resp.data.data.length 
     this.keep_cut_out_loss.push({
       result_cut_out_loss: this.result_cut_out_loss
     })

        worksheet.getCell("R"+ty).value = parseFloat(this.result_cut_out_loss).toFixed(2);
      worksheet.getCell("R"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
      this.result_cut_out_loss_per = 0;
      this.result_cut_out_loss_per = this.result_cut_out_loss / this.result_ttlmarker
      worksheet.getCell("S"+ty).value = parseFloat(this.result_cut_out_loss_per).toFixed(2);
      worksheet.getCell("S"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_remnants = 0
      for(var re_remnants = 0; re_remnants<resp.data.data.length; re_remnants++ ){
        this.result_remnants =  Number(this.result_remnants) + Number(this.rowexport[re_remnants].rement)
        
      }
     
     this.result_remnants = this.result_remnants/resp.data.data.length
     this.keep_remnants.push({
       result_remnants:this.result_remnants
     })

        worksheet.getCell("T"+ty).value = parseFloat(this.result_remnants).toFixed(2);
      worksheet.getCell("T"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.result_remnants_per = 0;
      this.result_remnants_per = this.result_remnants / this.result_ttlmarker
      worksheet.getCell("U"+ty).value = parseFloat(this.result_remnants_per).toFixed(2);
      worksheet.getCell("U"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


      this.result_total_length = 0
      for(var re_total_length = 0; re_total_length<resp.data.data.length; re_total_length++ ){
      this.result_total_length =  Number(this.result_total_length) + Number(this.rowexport[re_total_length].totallenthloss)
        
      }
     
     this.result_total_length = this.result_total_length/resp.data.data.length
     this.keep_total_loss.push({
       result_total_length:this.result_total_length
     }) 

        worksheet.getCell("V"+ty).value = parseFloat(this.result_total_length).toFixed(2);
      worksheet.getCell("V"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       this.result_total_length_per = 0;
      this.result_total_length_per = this.result_total_length / this.result_ttlmarker
      worksheet.getCell("W"+ty).value = parseFloat(this.result_total_length_per).toFixed(2);
      worksheet.getCell("W"+ty).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+ty).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };




       //B----------------------------------------------------------------
      worksheet.getCell("B"+a).value = "FL";
      worksheet.mergeCells("B"+a +":" + "B"+b);
      worksheet.getCell("B"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("B"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "B" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_table;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //C---------------------------------------------------------------------

       worksheet.getCell("C"+a).value = "S/O no.";
      worksheet.mergeCells("C"+a +":" + "C"+b);
      worksheet.getCell("C"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("C"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "C" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_no_short;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

           //D---------------------------------------------------------------------

       worksheet.getCell("D"+a).value = "Customer";
      worksheet.mergeCells("D"+a +":" + "D"+b);
      worksheet.getCell("D"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("D"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "D" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].customer;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

         //E---------------------------------------------------------------------------
       
       worksheet.getCell("E"+a).value = "GM08 no.";
      worksheet.mergeCells("E"+a +":" + "E"+b);
      worksheet.getCell("E"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("E"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "E" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].gm_no;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //F---------------------------------------------------------------
       if(this.multiple[x] == 1){
         this.name_f_type = "Woven"
       }
       if(this.multiple[x] == 2){
         this.name_f_type = "Light-Middle"
       }
      
      if(this.multiple[x] == 3){
         this.name_f_type = "Fleece"
       }
      if(this.multiple[x] == 4){
         this.name_f_type = "Special"
       }
    
      worksheet.getCell("F"+a).value = this.name_f_type;
      worksheet.getCell("F"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("F"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
           this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "F" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].fabric_type;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //G----------------------------------------------------------
       
       worksheet.getCell("G"+a).value = "Observe";
      worksheet.getCell("G"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("G"+b).value = "Width";
      worksheet.getCell("G"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("G"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "G" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].obs_width;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //H----------------------------------------------------------
        worksheet.getCell("H"+a).value = "Width";
      worksheet.getCell("H"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("H"+b).value = "(inch)";
      worksheet.getCell("H"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("H"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "H" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].width_inch;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }
       //I--------------------------------------------------------------
         worksheet.getCell("I"+a).value = "Length";
      worksheet.getCell("I"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("I"+b).value = "(yd)";
      worksheet.getCell("I"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("I"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
      this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "I" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].length_ydb;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

         //J-----------------------------------------------------------
         worksheet.getCell("J"+a).value = "Ttl Marker";
      worksheet.getCell("J"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("J"+b).value = "Length(yd)";
      worksheet.getCell("J"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "J" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].ttl_marker;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
       }; 
       }

        //K-------------------------------------------------------------------------------
         worksheet.getCell("K"+a).value = "Plug1+2";
      worksheet.getCell("K"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("K"+b).value = "(inch)";
      worksheet.getCell("K"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("K"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
         
         this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "K" + ab
      
            var z = ab-a-2;
      worksheet.getCell(y).value = parseFloat(check_null(this.rowexport[z].plug12)).toFixed(2);
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
       }; 
       }

       //L------------------------------------------------------------------
            worksheet.getCell("L"+a).value = "%";
      worksheet.getCell("L"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("L"+b).value = "1.50%";
      worksheet.getCell("L"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

          this.length =  resp.data.data.length+b //2 มาจากไหน?
       
       for(var ab=a+2; ab<=this.length; ab++){
      var y  = "L" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].widthloss;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
       }; 
       }

        //M--------------------------------------------------------------
        worksheet.getCell("M"+a).value = "Plug1+2";
      worksheet.getCell("M"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("M"+b).value = "(inch)";
      worksheet.getCell("M"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "M" + ab
      
            var z = ab-a-2;

      this.result_plug12_inch = parseFloat((this.rowexport[z].endless_length_yd*36)/(this.rowexport[z].ttl_marker/this.rowexport[z].length_ydb)).toFixed(2)
       worksheet.getCell(y).value = this.result_plug12_inch;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


      //N----------------------------------------------------------
       worksheet.getCell("N"+a).value = "Length";
      worksheet.getCell("N"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("N"+b).value = "(yd)";
      worksheet.getCell("N"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "N" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].endless_length_yd;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }


       //O------------------------------------------------------------------------
        worksheet.getCell("O"+a).value = "%";
      worksheet.getCell("O"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("O"+b).value = "0.15%";
      worksheet.getCell("O"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "O" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].avg_end+"%";
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //P--------------------------------------------------------------------------
         worksheet.getCell("P"+a).value = "Length";
      worksheet.getCell("P"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("P"+b).value = "(yd)";
      worksheet.getCell("P"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "P" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].spliceloss;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //Q------------------------------------------------------------------------------
         worksheet.getCell("Q"+a).value = "%";
      worksheet.getCell("Q"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("Q"+b).value = "0.20%";
      worksheet.getCell("Q"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "Q" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].splicelossper+"%";
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //R------------------------------------------------------------------------------
         worksheet.getCell("R"+a).value = "Length";
      worksheet.getCell("R"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("R"+b).value = "(yd)";
      worksheet.getCell("R"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
         this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "R" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].totalcutout;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //S------------------------------------------------------------------------------
         worksheet.getCell("S"+a).value = "%";
      worksheet.getCell("S"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("S"+b).value = "0.00%";
      worksheet.getCell("S"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "S" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = parseFloat(this.rowexport[z].totalcutoutper).toFixed(2)+"%";
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      //T-------------------------------------------------------------------------------

         worksheet.getCell("T"+a).value = "Length";
      worksheet.getCell("T"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("T"+b).value = "(yd)";
      worksheet.getCell("T"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "T" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].rement;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //U-------------------------------------------------------------------------------

         worksheet.getCell("U"+a).value = "%";
      worksheet.getCell("U"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("U"+b).value = "0.14%";
      worksheet.getCell("U"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "U" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].rement_per+"%";
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //V-------------------------------------------------------------------------------

         worksheet.getCell("V"+a).value = "Length";
      worksheet.getCell("V"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("V"+b).value = "(yd)";
      worksheet.getCell("V"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "V" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].totallenthloss;
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

       //W-------------------------------------------------------------------------------

         worksheet.getCell("W"+a).value = "%of";
      worksheet.getCell("W"+a).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+a).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       worksheet.getCell("W"+b).value = "0.49%";
      worksheet.getCell("W"+b).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("W"+b).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.length =  resp.data.data.length+a+2
       
       for(var ab=a+2; ab<this.length; ab++){
      var y  = "W" + ab
      
            var z = ab-a-2;
       worksheet.getCell(y).value = this.rowexport[z].totallenthlossper+"%";
      worksheet.getCell(y).alignment = { horizontal: "center" };
      worksheet.getCell(y).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
       }

      
         }  
         

      
         count++
          }else{
          this.$q.notify({
            message: this.multiple[x]+": Don't Have Data",

            color: "red",
          });
        }
         console.log(this.multiple.length)
         console.log(count-1)
         if(this.multiple.length == count){
          
      var c = this.length+2
      var d = this.length+3     
      var e = this.length+5
      var f = this.length+6
      var  g = this.length+7
      var  h = this.length+8
      var i = this.length+9
      var j = this.length+10
      var k = this.length+11
      var l = this.length+12 //endloss
      var m = this.length+13 //splice loss
      var n = this.length+14 //Cut - out Loss
      var o = this.length+15 //Remnants Loss
      var p = this.length+16 //Total Loss

      worksheet.getCell("A"+c).value = "Total Target Weight";
      worksheet.mergeCells("A"+c+":" +"I"+c);
      worksheet.getCell("A"+c).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("J"+c).value = "";
      worksheet.getCell("J"+c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("K"+c).value = "";
      worksheet.getCell("K"+c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("L"+c).value = "1.50%";
      worksheet.getCell("L"+c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("M"+c).value = "";
      worksheet.getCell("M"+c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        worksheet.getCell("N"+c).value = "";
      worksheet.getCell("N"+c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("O"+c).value = "0.17%";
      worksheet.getCell("O"+c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("P"+c).value = "";
      worksheet.getCell("P"+c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("Q"+c).value = "0.20%";
      worksheet.getCell("Q"+c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("R"+c).value = "";
      worksheet.getCell("R"+c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("S"+c).value = "0.00%";
      worksheet.getCell("S"+c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        worksheet.getCell("T"+c).value = "";
      worksheet.getCell("T"+c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("U"+c).value = "0.16%";
      worksheet.getCell("U"+c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        worksheet.getCell("V"+c).value = "";
      worksheet.getCell("V"+c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("W"+c).value = "0.53%";
      worksheet.getCell("W"+c).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      let org = this.$q.localStorage.getItem("org");
     
      worksheet.getCell("A"+d).value = "NY"+org+ " : Daily Total";
      worksheet.mergeCells("A"+d+":" +"I"+d);
      worksheet.getCell("A"+d).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("A"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.result_sum_ttlmarker = 0
      for(var jd = 0; jd<this.keep_ttlmarker.length; jd++){
        this.result_sum_ttlmarker = Number(this.result_sum_ttlmarker)+Number(this.keep_ttlmarker[jd].result_ttlmarker)
       
      }
      
      worksheet.getCell("J"+d).value = this.result_sum_ttlmarker;
      worksheet.getCell("J"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
   

        this.sum_plug12_inch = 0
        this.result_sum_plug12_inch = 0
       
      for(var kd = 0; kd<this.keep_ttlmarker.length; kd++){
         this.result_sum_plug12_inch = this.keep_ttlmarker[kd].result_ttlmarker * this.keep_plug12_inch[kd].result_plug12_inch
        

  
        this.sum_plug12_inch = Number(this.sum_plug12_inch) + Number(this.result_sum_plug12_inch)
      }
     
      
      console.log(this.sum_plug12_inch)
       worksheet.getCell("K"+d).value =  parseFloat(this.sum_plug12_inch).toFixed(3);
      worksheet.getCell("K"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.result_sum_plug12_inch_l=0
      this.sum_plug12_inch_l =0
       for(var ld = 0; ld<this.keep_ttlmarker.length; ld++){
         this.result_sum_plug12_inch_l = this.keep_ttlmarker[ld].result_ttlmarker * this.keep_perplug12[ld].result_perplug12
        

  
        this.sum_plug12_inch_l = Number(this.sum_plug12_inch) + Number(this.result_sum_plug12_inch)
      }
      this.sum_plug12_inch_l = this.sum_plug12_inch_l/this.result_sum_ttlmarker
      
      worksheet.getCell("L"+d).value = this.sum_plug12_inch_l;
      worksheet.getCell("L"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

  
          this.result_avg_end_plug12 = 0
          this.sum_avg_end_plug12 = 0
         for(var md = 0; md<this.keep_ttlmarker.length; md++){
          
         this.result_avg_end_plug12 = this.keep_ttlmarker[md].result_ttlmarker * this.keep_avg_loss_plug12.result_avg_end_plug12
        

  
        this.sum_avg_end_plug12 = Number(this.sum_avg_end_plug12) + Number(this.result_sum_plug12_inch)
        
      } 
     this.sum_avg_end_plug12 = this.sum_avg_end_plug12/this.result_sum_ttlmarker
      worksheet.getCell("M"+d).value = parseFloat(this.sum_avg_end_plug12).toFixed(3);
      worksheet.getCell("M"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
  
      worksheet.getCell("N"+d).value = parseFloat(this.sum_length_yd).toFixed(3);
      worksheet.getCell("N"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
      this.sum_avg_per = this.sum_length_yd/this.result_sum_ttlmarker
  
       worksheet.getCell("O"+d).value = parseFloat(this.sum_avg_per).toFixed(3);
      worksheet.getCell("O"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
       worksheet.getCell("P"+d).value = parseFloat(this.sum_result_splice).toFixed(3);
      worksheet.getCell("P"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.sum_result_splice_per = this.sum_result_splice/this.result_sum_ttlmarker
      
       worksheet.getCell("Q"+d).value = parseFloat(this.sum_result_splice_per).toFixed(3);
      worksheet.getCell("Q"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
       worksheet.getCell("R"+d).value =  parseFloat(this.sum_cut_loss_out).toFixed(3);
      worksheet.getCell("R"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
     
      this.sum_cut_loss_out_per = this.sum_cut_loss_out/this.result_sum_ttlmarker
     
       worksheet.getCell("S"+d).value =  parseFloat(this.sum_cut_loss_out_per).toFixed(3);
      worksheet.getCell("S"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("T"+d).value = parseFloat(this.sum_remnants).toFixed(3);
      worksheet.getCell("T"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.sum_remnants_per = this.sum_remnants/this.result_sum_ttlmarker
       worksheet.getCell("U"+d).value = parseFloat(this.sum_remnants_per).toFixed(3);
      worksheet.getCell("U"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("V"+d).value = parseFloat(this.sum_total_length).toFixed(3);
      worksheet.getCell("V"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.sum_total_length_per = this.sum_total_length/this.result_sum_ttlmarker
       worksheet.getCell("W"+d).value = parseFloat(this.sum_total_length_per).toFixed(3);
      worksheet.getCell("W"+d).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };










      
      
         worksheet.getCell("L"+e).value = "หมายเหตุ  NYG1 :  % Spread Loss โดยรวม = 0.46% ต่ำกว่า Target  (0.53%) ความสูญเสียเกิดจาก";
      
      worksheet.getCell("L"+e).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 
        
       worksheet.getCell("N"+f).value = "% End Loss = 0.37% (Target 0.17%)  ซึ่งเป็นจุดที่ควรควบคุม";
    
      worksheet.getCell("N"+f).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
       worksheet.getCell("N"+g).value = "% Width Loss  โดยรวม =2.08%  สูงกว่า Target (1.50%) ";
      worksheet.getCell("N"+g).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
       worksheet.getCell("L"+h).value = "Target";
      worksheet.mergeCells("L"+h+":" +"O"+h);
      worksheet.getCell("L"+h).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+h).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       
       worksheet.getCell("P"+h).value = "Weighted target of this Week";
      worksheet.mergeCells("P"+h+":" +"P"+i);
      worksheet.getCell("P"+h).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+h).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 

        worksheet.getCell("Q"+h).value = "Actual Performance";
      worksheet.mergeCells("Q"+h+":" +"V"+h);
      worksheet.getCell("Q"+h).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+h).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("L"+i).value = "Woven";
      worksheet.mergeCells("L"+i+":" +"L"+j);
      worksheet.getCell("L"+i).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+i).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("Q"+i).value = "Woven";
      worksheet.mergeCells("Q"+i+":" +"Q"+j);
      worksheet.getCell("Q"+i).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+i).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("M"+i).value = "Knit";
      worksheet.mergeCells("M"+i+":" +"O"+i);
      worksheet.getCell("M"+i).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+i).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
        worksheet.getCell("R"+i).value = "Knit";
      worksheet.mergeCells("R"+i+":" +"T"+i);
      worksheet.getCell("R"+i).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+i).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      }; 


        worksheet.getCell("U"+i).value = "of This Week";
      worksheet.mergeCells("U"+i+":" +"V"+i);
      worksheet.getCell("U"+i).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+i).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("M"+j).value = "Light & Mid Wt.";
      worksheet.getCell("M"+j).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("N"+j).value = "Fleece";
      worksheet.getCell("N"+j).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("O"+j).value = "Special";
      worksheet.getCell("O"+j).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("P"+j).value = "Overall";
      worksheet.getCell("P"+j).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("P"+j).value = "Overall";
      worksheet.getCell("P"+j).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("R"+j).value = "Light & Mid Wt.";
      worksheet.getCell("R"+j).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("S"+j).value = "Fleece";
      worksheet.getCell("S"+j).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("T"+j).value = "Special";
      worksheet.getCell("T"+j).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("U"+j).value = "Average";
      worksheet.getCell("U"+j).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("V"+j).value = "Total";
      worksheet.getCell("V"+j).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
      

        worksheet.getCell("L"+k).value = "";
      worksheet.getCell("L"+k).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
        worksheet.getCell("M"+k).value = "";
      worksheet.getCell("M"+k).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };


      worksheet.getCell("N"+k).value = "";
      worksheet.getCell("N"+k).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("O"+k).value = "";
      worksheet.getCell("O"+k).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("P"+k).value = "";
      worksheet.getCell("P"+k).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("Q"+k).value = "";
      worksheet.getCell("Q"+k).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

       worksheet.getCell("R"+k).value = "";
      worksheet.getCell("R"+k).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("S"+k).value = "";
      worksheet.getCell("S"+k).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("T"+k).value = "";
      worksheet.getCell("T"+k).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       
      
      this.devide_ttlmarker = this.result_sum_ttlmarker/this.keep_ttlmarker.length
         worksheet.getCell("U"+k).value = parseFloat(this.devide_ttlmarker).toFixed(3);
      worksheet.getCell("U"+k).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    
      worksheet.getCell("V"+k).value = this.result_sum_ttlmarker;
      worksheet.getCell("V"+k).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      
         worksheet.getCell("J"+l).value = "End Loss";
      worksheet.mergeCells("J"+l+":" +"K"+l);
      worksheet.getCell("J"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        worksheet.getCell("L"+l).value = "0.23%";
      worksheet.getCell("L"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("M"+l).value = "0.32%";
      worksheet.getCell("M"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }, 
      };

      worksheet.getCell("N"+l).value = "0.35%";
      worksheet.getCell("N"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("O"+l).value = "0.45%";
      worksheet.getCell("O"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.sum_x_ttlmarker = 0
      this.value_ttlmarker = ['0.23','0.32','0.35','0.45']
      for(var pl = 0; pl<this.keep_ttlmarker.length; pl++){
          this.sum_x_ttlmarker = this.value_ttlmarker[pl]*this.keep_ttlmarker[pl].result_ttlmarker
         console.log(this.sum_x_ttlmarker)
      }
      this.sum_x_ttlmarkerx = this.sum_x_ttlmarker/this.result_sum_ttlmarker
      
      worksheet.getCell("P"+l).value = this.sum_x_ttlmarkerx;
      worksheet.getCell("P"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      if(this.keep_avg_end.length>0){
      worksheet.getCell("Q"+l).value = this.keep_avg_end[0].result_avg_per
      worksheet.getCell("Q"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }else{
         worksheet.getCell("Q"+l).value = 0
      worksheet.getCell("Q"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }
      if(this.keep_avg_end.length>1){
      
       worksheet.getCell("R"+l).value = this.keep_avg_end[1].result_avg_per
      worksheet.getCell("R"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      }else{

         worksheet.getCell("R"+l).value = 0
      worksheet.getCell("R"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }
      
      if(this.keep_avg_end.length>2){
       worksheet.getCell("S"+l).value = parseFloat(this.keep_avg_end[2].result_avg_per).toFixed(5);
      worksheet.getCell("S"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }else{
         worksheet.getCell("S"+l).value = 0
      worksheet.getCell("S"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }

      if(this.keep_avg_end.length>3){
        worksheet.getCell("T"+l).value = this.keep_avg_end[3].result_avg_per;
      worksheet.getCell("T"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        
         }else{

      worksheet.getCell("T"+l).value = 0
      worksheet.getCell("T"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

         }

      worksheet.getCell("U"+l).value = "";
      worksheet.getCell("U"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("V"+l).value = "";
      worksheet.getCell("V"+l).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+l).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      


        worksheet.getCell("J"+m).value = "Splice Loss";
      worksheet.mergeCells("J"+m+":" +"K"+m);
      worksheet.getCell("J"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        worksheet.getCell("L"+m).value = "0.10%";
      worksheet.getCell("L"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("M"+m).value = "0.10%";
      worksheet.getCell("M"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("N"+m).value = "0.10%";
      worksheet.getCell("N"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("O"+m).value = "0.10%";
      worksheet.getCell("O"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       this.sum_x_splice_loss = 0
      this.value_splice_loss = ['0.10','0.10','0.10','0.10']
      for(var pm = 0; pm<this.keep_ttlmarker.length; pm++){
         this.sum_x_splice_loss = this.value_splice_loss[pm]*this.keep_ttlmarker[pm].result_ttlmarker
      }
      this.sum_x_splice_lossx = this.sum_x_splice_loss/this.result_sum_ttlmarker

      worksheet.getCell("P"+m).value = this.sum_x_splice_lossx
      worksheet.getCell("P"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
      if(this.keep_splice_Loss.length>0){
      worksheet.getCell("Q"+m).value = this.keep_splice_Loss[0].result_splice
      worksheet.getCell("Q"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }else{
        worksheet.getCell("Q"+m).value = 0
      worksheet.getCell("Q"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }

      if(this.keep_splice_Loss.length>1){
      
       worksheet.getCell("R"+m).value = this.keep_splice_Loss[1].result_splice
      worksheet.getCell("R"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }else{
         worksheet.getCell("R"+m).value = 0
      worksheet.getCell("R"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }

      if(this.keep_splice_Loss.length>2){
     
       worksheet.getCell("S"+m).value = parseFloat(this.keep_splice_Loss[2].result_splice).toFixed(5)
      worksheet.getCell("S"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }else{
 worksheet.getCell("S"+m).value = 0
      worksheet.getCell("S"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }

      if(this.keep_splice_Loss.length>3){
      
        worksheet.getCell("T"+m).value = this.keep_splice_Loss[3].result_splice;
      worksheet.getCell("T"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }else{
  worksheet.getCell("T"+m).value = 0
      worksheet.getCell("T"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }
      worksheet.getCell("U"+m).value = "";
      worksheet.getCell("U"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("V"+m).value = "";
      worksheet.getCell("V"+m).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+m).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      

        worksheet.getCell("J"+n).value = "Cut - Out Loss";
      worksheet.mergeCells("J"+n+":" +"K"+n);
      worksheet.getCell("J"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("L"+n).value = "0.00%";
      worksheet.getCell("L"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("M"+n).value = "0.00%";
      worksheet.getCell("M"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("N"+n).value = "0.00%";
      worksheet.getCell("N"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("O"+n).value = "0.00%";
      worksheet.getCell("O"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        this.sum_cut_out_loss = 0
      this.value_cut_out_loss = ['0.00','0.00','0.00','0.00']
      for(var pn = 0; pn<this.keep_ttlmarker.length; pn++){
         this.sum_cut_out_loss = this.value_cut_out_loss[pn]*this.keep_ttlmarker[pn].result_ttlmarker
      }
      this.sum_cut_out_lossx = this.sum_cut_out_loss/this.result_sum_ttlmarker
      worksheet.getCell("P"+n).value = this.sum_cut_out_lossx
      worksheet.getCell("P"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      if(this.keep_cut_out_loss>0){
      worksheet.getCell("Q"+n).value = this.keep_cut_out_loss[0].result_cut_out_loss;
      worksheet.getCell("Q"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }else{
        worksheet.getCell("Q"+n).value = 0;
      worksheet.getCell("Q"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }

       if(this.keep_cut_out_loss>1){
      
       worksheet.getCell("R"+n).value = this.keep_cut_out_loss[1].result_cut_out_loss
      worksheet.getCell("R"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       }else{
      worksheet.getCell("R"+n).value = 0
      worksheet.getCell("R"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      }

      if(this.keep_cut_out_loss>2){
      
       worksheet.getCell("S"+n).value = parseFloat(this.keep_cut_out_loss[2].result_cut_out_loss).toFixed(5)
      worksheet.getCell("S"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      }else{
       worksheet.getCell("S"+n).value = 0
      worksheet.getCell("S"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      }
      if(this.keep_cut_out_loss>3){
      
        worksheet.getCell("T"+n).value = this.keep_cut_out_loss[3].result_cut_out_loss
      worksheet.getCell("T"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }else{
        worksheet.getCell("T"+n).value = 0
        worksheet.getCell("T"+n).alignment = { horizontal: "center",vertical:"middle" };
        worksheet.getCell("T"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      }

      worksheet.getCell("U"+n).value = "";
      worksheet.getCell("U"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("V"+n).value = "";
      worksheet.getCell("V"+n).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+n).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      

        worksheet.getCell("J"+o).value = "Remnants Loss";
      worksheet.mergeCells("J"+o+":" +"K"+o);
      worksheet.getCell("J"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        worksheet.getCell("L"+o).value = "0.05%";
      worksheet.getCell("L"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("M"+o).value = "0.08%";
      worksheet.getCell("M"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("N"+o).value = "0.11%";
      worksheet.getCell("N"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("O"+o).value = "0.10%";
      worksheet.getCell("O"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      this.sum_remnants_loss = 0
      this.value_remnants_loss = ['0.05','0.08','0.11','0.10']
      for(var po = 0; po<this.keep_ttlmarker.length; po++){
         this.sum_remnants_loss = this.value_cut_out_loss[po]*this.keep_ttlmarker[po].result_ttlmarker
      }
      this.sum_remnants_lossx = this.sum_remnants_loss/this.result_sum_ttlmarker
      worksheet.getCell("P"+o).value = this.sum_remnants_lossx
      worksheet.getCell("P"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      if(this.keep_remnants.length>0){
      worksheet.getCell("Q"+o).value = this.keep_remnants[0].result_remnants
      worksheet.getCell("Q"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }else{
          worksheet.getCell("Q"+o).value = 0
      worksheet.getCell("Q"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }

       if(this.keep_remnants.length>1){
       worksheet.getCell("R"+o).value = this.keep_remnants[1].result_remnants
      worksheet.getCell("R"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       }else{
         worksheet.getCell("R"+o).value = 0
      worksheet.getCell("R"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };  

       }

      if(this.keep_remnants.length>2){
       worksheet.getCell("S"+o).value =  parseFloat(this.keep_remnants[2].result_remnants).toFixed(5) 
      worksheet.getCell("S"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }else{
      worksheet.getCell("S"+o).value = 0
      worksheet.getCell("S"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      }

          if(this.keep_remnants.length>3){
     
        worksheet.getCell("T"+o).value =  this.keep_remnants[3].result_remnants
      worksheet.getCell("T"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
          }else{
      worksheet.getCell("T"+o).value =  0
      worksheet.getCell("T"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

          }
      worksheet.getCell("U"+o).value = "";
      worksheet.getCell("U"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("V"+o).value = "";
      worksheet.getCell("V"+o).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+o).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
     
        worksheet.getCell("J"+p).value = "Total Loss";
      worksheet.mergeCells("J"+p+":" +"K"+p);
      worksheet.getCell("J"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("J"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
        worksheet.getCell("L"+p).value = "0.05%";
      worksheet.getCell("L"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("L"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

        worksheet.getCell("M"+p).value = "0.08%";
      worksheet.getCell("M"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("M"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("N"+p).value = "0.11%";
      worksheet.getCell("N"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("N"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("O"+p).value = "0.10%";
      worksheet.getCell("O"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("O"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      this.sum_total_loss = 0
      this.value_sum_total_loss = ['0.05','0.08','0.11','0.10']
      for(var pp = 0; pp<this.keep_ttlmarker.length; pp++){
         this.sum_total_loss = this.value_sum_total_loss[pp]*this.keep_ttlmarker[pp].result_ttlmarker
      }
      this.sum_total_lossx = this.sum_total_loss/this.result_sum_ttlmarker 
      worksheet.getCell("P"+p).value = this.sum_total_lossx;
      worksheet.getCell("P"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("P"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      
      if(this.keep_total_loss.length>0){
      worksheet.getCell("Q"+p).value =  parseFloat(this.keep_total_loss[0].result_total_length).toFixed(4) 
      worksheet.getCell("Q"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }else{
      worksheet.getCell("Q"+p).value = 0
      worksheet.getCell("Q"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("Q"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      }

      if(this.keep_total_loss.length>1){
       worksheet.getCell("R"+p).value =  parseFloat(this.keep_total_loss[1].result_total_length).toFixed(4) 
      worksheet.getCell("R"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      }else{
         worksheet.getCell("R"+p).value = 0
      worksheet.getCell("R"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("R"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       }

       if(this.keep_total_loss.length>2){
       worksheet.getCell("S"+p).value =  parseFloat(this.keep_total_loss[2].result_total_length).toFixed(4);
      worksheet.getCell("S"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       }else{
           worksheet.getCell("S"+p).value = 0
      worksheet.getCell("S"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("S"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       }
       if(this.keep_total_loss.length>3){
        worksheet.getCell("T"+p).value =  parseFloat(this.keep_total_loss[3].result_total_length).toFixed(4) 
      worksheet.getCell("T"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       }else{
          worksheet.getCell("T"+p).value = 0
      worksheet.getCell("T"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("T"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
       }

      worksheet.getCell("U"+p).value = "";
      worksheet.getCell("U"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("U"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };

      worksheet.getCell("V"+p).value = "";
      worksheet.getCell("V"+p).alignment = { horizontal: "center",vertical:"middle" };
      worksheet.getCell("V"+p).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      


        workbook.xlsx.writeBuffer().then((data) => {
        const blob = new Blob([data], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8",
        });
        saveAs(blob, "report.xlsx");
      });
         }
  
      
      

</script>

<style>

</style>