
 
alert('hi');
  var input = $('#url');
  var xlsxUrl = '';
  var admJson = $('#admjson');
  var convertBtn = $('#convert-btn');
  var preview = $('#preview');
  var convertingContainer = $('#converting');
  var fileTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
  var converting = false;
  function returnFileSize(number) {
  if(number < 1024) {
  return number + 'bytes';
  } else if(number >= 1024 && number < 1048576) {
  return (number/1024).toFixed(1) + 'KB';
  } else if(number >= 1048576) {
  return (number/1048576).toFixed(1) + 'MB';
  }
  }
  function validFileType(file) {
    for(var i = 0; i < fileTypes.length; i++) {
      if(file.type === fileTypes[i]) {
        convertBtn.disabled = false;
        var p = '<p>File name ' + file.name + ', file size ' + returnFileSize(file.size) + '.</p>';
        preview.append(p);
        return true;
      } else {
        converting.style.display = "none";
        var p = 'File name ' + file.name + ': Not a valid file type. Update your selection.';
        preview.append(p);      
      }
   }     
   return false;
  }
    
convertBtn.click(function () {
  if (converting) return;
  xlsxUrl = input.val();
  if (xlsxUrl.length == 0)  {    
    alert('Please enter xlsx url');
  } else {        
      convertBtn.prop('disabled', true);        
      converting = true;
      admJson.text('');   
      convertingContainer.show();
      preview.html('');
      $.ajaxSetup({
        beforeSend:function(jqXHR,settings){
          if (settings.dataType === 'binary'){
            settings.xhr().responseType='arraybuffer';
            settings.processData=false;
          }
        }
      })
      
      //use ajax now
      $.ajax({
        url: xlsxUrl,
        dataType:"binary"
      }).done(function(data) {
      // If successful
        
        /* convert data to binary string */
        var data = new Uint8Array(data);
        var arr = new Array();
        for (var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
        var bstr = arr.join("");
        
        /* Call XLSX */
        var workbook = XLSX.read(bstr, {
        type: "binary"
        });
        
        /* DO SOMETHING WITH workbook HERE */
        var first_sheet_name = workbook.SheetNames[0];
        /* Get worksheet */
        var worksheet = workbook.Sheets[first_sheet_name];
        jsonVar = XLSX.utils.sheet_to_json(worksheet, {raw: true});
        jsonStr = JSON.stringify(jsonVar);
        admJson.text(jsonStr);
        converting = false;
        convertingContainer.hide();
        convertBtn.prop('disabled', false);
      }).fail(function(jqXHR, textStatus, errorThrown) {
      // If fail
        converting = false;
        convertingContainer.hide();
        convertBtn.prop('disabled', false);           
        console.log(textStatus + ': ' + errorThrown);
        preview.html('<p>Error, cannot load xlsx file: ' + xlsxUrl + '. Error detail: ' + textStatus + ': ' + errorThrown + '</p>');
      });                   
   }
  });


