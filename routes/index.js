var express = require('express')
var router = express.Router()
var multer = require('multer');
var xlsxj = require("xlsx-to-json");
var _ = require('underscore');
var json2xls = require('json2xls');
var fs = require('fs');
var path = require('path');
var xl = require('excel4node');


var storage = multer.diskStorage({

  destination: function (req, file, cb) {
    cb(null, 'imports')
  },

  filename: function (req, file, cb) {
    cb(null, 'potukuru_input.xlsx')
  }
})
 
var upload = multer({ storage: storage })

/* GET */
router.get('/', function(req, res, next) {
  res.render('index', { success: false })
})

/* POST */
router.post('/', upload.single('inputfile'),(req,res) => {
  
  xlsxj({
    input: "imports/potukuru_input.xlsx",
    output: null
  }, 
  
  function(err, result) {
    if(err) {
      console.error(err)
    }else {
      console.log(result)

    var sorteddata =  _.sortBy( result, function( item ) { return -item['Critic Score'] && item['Genre'] } )

    var excelborder =  { 
      left: {
          style: 'thin', 
          color: '000000' 
      },
      right: {
          style: 'thin',
          color: '000000'
      },
      top: {
          style: 'thin',
          color: '000000'
      },
      bottom: {
          style: 'thin',
          color: '000000'
      }
    }
    
    var workbook = new xl.Workbook({
       defaultFont: {
      size: 11,
      name: 'Calibri',
      color: 'FF000000'
      }
     })
    
    
     var addWorksheet = workbook.addWorksheet('Output');     
   
     var outputstyle = workbook.createStyle({
    
      font: {
           bold: true,
           color: '000000'
       },
       fill: {
             type: 'pattern',
             patternType: 'solid',
             fgColor: 'C6E0B4' 
         },
         border: excelborder
   })

        addWorksheet.cell(2,2,2,4).string('Name').style(outputstyle)

        var outputstyle1 = workbook.createStyle({
          font:{
            underline: true,
            italics: true,
            bold: false
          },
          border: excelborder
        })
       
        addWorksheet.cell(2,3,2,4,true).string('Potukuru,Durga Charan').style(outputstyle1)
        
        var outputstyle2 = workbook.createStyle({
          font:{
            bold: true,
            color: 'FFFFFF'
          },
           fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: 'C00000' 
        },
        border: excelborder,
        alignment:{
          horizontal: 'center'
        }
        })

        var p = 1;
        _.each(['SNO','Genre','Credit Score','Album Name','Artist','Release Date'], function(ele){
          addWorksheet.cell(4,p).string(ele).style(outputstyle2)
          p++;
        })


        p = 5;
      
        var outputcolor = 'FFF2CC'
        var outputcolor1 = 'C6E0B4'
        var pastgenre = sorteddata[0]['Genre']

     _.each(sorteddata, function(ele){
        ele['Credit Score'] = ele['Critic Score']
        delete ele['Critic Score']
        q = 1;
        if(pastgenre != ele['Genre']){
          temp = outputcolor;
          outputcolor = outputcolor1;
          outputcolor1 = temp;
          pastgenre = ele['Genre']
        }
        _.each(['SNO','Genre','Credit Score','Album Name','Artist','Release Date'], function(attr){
          var outputbodystyle = workbook.createStyle({
            fill: {
             type: 'pattern',
             patternType: 'solid',
             fgColor: outputcolor 
            },
            border: excelborder,
            alignment: {
              horizontal: /\d/.test(ele[attr]) ? 'right' : 'left'
            }
         })
          addWorksheet.cell(p,q).string(ele[attr]).style(outputbodystyle)
          q++;
        })
        p++;
        
      })
        workbook.write('./public/potukuru_output.xlsx', function(err,stats){
          if(err) return res.send('Error')
            res.download(path.join(__dirname, '../public/potukuru_output.xlsx'),'potukuru_output.xlsx')
        });
       
    }
  })
})
module.exports = router