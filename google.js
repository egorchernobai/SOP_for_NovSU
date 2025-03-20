function doPost(e) {
    try {
      Logger.log("Получены данные: " + e.postData.contents); // Логируем входящие данные
      var data = JSON.parse(e.postData.contents);
  
      if (!data.formTitle || !data.json1) {
        return ContentService.createTextOutput("Ошибка: Отсутствуют formTitle или json1");
      }
  
      var formUrl = createGoogleForm(data);
      return ContentService.createTextOutput(formUrl);
    } catch (error) {
      Logger.log("Ошибка: " + error.toString());
      return ContentService.createTextOutput("Ошибка: " + error.toString());
    }
  }
  
  function createGoogleForm(data) {
    var form = FormApp.create(data.formTitle);
    form.setDescription(data.description);
    
    try {
      for (var key in data.json1) { 
        if (data.json1.hasOwnProperty(key)) {
          form.addPageBreakItem().setTitle(key).setHelpText("Оценка дисциплины");
          subject(form, data.json2, data.json1[key], key);
  
          var teachers = data.json1[key];
          for (var j = 0; j < teachers.length; j++) {
            form.addPageBreakItem().setTitle(teachers[j]).setHelpText("Оценка преподавателя");
            teacher(form, data.json3);
          }
        }
      }
      return form.getEditUrl();
    } catch (error) {
      return "Ошибка: " + error.toString();
    }
  }
  
  function subject(form, json2, teachers, key){
    if (!json2) return; // Проверяем, что json2 не пуст
  
    for (var quest in json2){
      if (json2.hasOwnProperty(quest) && Array.isArray(json2[quest]) && json2[quest].length > 1){
        var item = form.addMultipleChoiceItem().setTitle(quest);
        var otvets = json2[quest].slice(0, -1); // Берем все, кроме последнего элемента
        if(json2[quest][json2[quest].length-1]==1){
          item.setRequired(true);
        }
  
        item.setChoiceValues(otvets);
      } else if(json2.hasOwnProperty(quest) && Array.isArray(json2[quest])){
        var item = form.addTextItem().setTitle(quest);
        if(json2[quest][json2[quest].length-1]==1){
          item.setRequired(true);
        }
      }
    }
  }

  function teacher(form, json2){
    if (!json2) return; // Проверяем, что json2 не пуст
  
    for (var quest in json2){
      if (json2.hasOwnProperty(quest) && Array.isArray(json2[quest]) && json2[quest].length > 1){
        var item = form.addMultipleChoiceItem().setTitle(quest);
        var otvets = json2[quest].slice(0, -1); // Берем все, кроме последнего элемента
        if(json2[quest][json2[quest].length-1]==1){
          item.setRequired(true);
        }
  
        item.setChoiceValues(otvets);
      } else if(json2.hasOwnProperty(quest) && Array.isArray(json2[quest])){
        var item = form.addTextItem().setTitle(quest);
        if(json2[quest][json2[quest].length-1]==1){
          item.setRequired(true);
        }
      }
    }
  }  
  