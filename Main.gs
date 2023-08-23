////////////////////////////////////////////////////////////////////////
//                                                                    //
// Script Developer: Kayla Tinker (WFO Anchorage)                     //
//                                                                    //
// Some additional coding from CR IDSS Creators                       //
//                                                                    //
//////////////////////////////////////////////////////////////////////// 
// EDITS - changed 0 to 1 to grab old warning

//
// Config if needed
//
var timeZone = "America/Anchorage" // ALASKA use Anchorage...PACIFIC use Vancouver...MOUNTAIN use Denver...CENTRAL use Chicago...EASTERN use New_York

//
// Creates menu for user
//
function onOpen() {

  SlidesApp.getUi()
    .createMenu('Run Script')
    .addItem('Update Date Stamp','updateDates')
    .addItem('Update Text','updateText')
    .addToUi(); 
}

//
// Updates the date stamp text box to the current time
//
function updateDates() {

  var ui = SlidesApp.getUi(); 
  var response = ui.alert('The selected timezone is '+timeZone+'. Is that your current timezone?', ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    ui.alert('Date/Time will be represented in the ' + timeZone + ' timezone.');
  } else if (response == ui.Button.NO) {
    ui.alert('Edit script for the appropriate timezone. Tools -> Script -> adjust line 12.');
  };
  
  var title = "DateStamp"; 
  var date = Utilities.formatDate(new Date(), timeZone,"MM/dd/yyyy h:mm a z");
  var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();
  //getSlides()[0];
  var date = 'Issued: '+date;
  var shapes = slide.getShapes();
  var s = shapes.filter(function(e) {return e.getTitle() == title});
  if (s.length > 0) {
    s[0].getText().setText(date);
  } else {
    var shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 450, 55, 300, 60).setTitle(title);
    var textRange = shape.getText();
    textRange.setText(date);
    textRange.getTextStyle().setFontSize(16).setForegroundColor('#ffffff');
  }

};

//
// Update the text from wproduct API
//
function updateText() {
  var ui = SlidesApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Enter one zone in Product' 
      + '\r\n(ex: 17)',
      'Zone:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var zone = result.getResponseText();
  if (button == ui.Button.OK) {
  // User clicked "OK".
    ui.alert('One zone in this product is zone ' + zone + '.');
  } else if (button == ui.Button.CANCEL) {
  // User clicked "Cancel".
    ui.alert('No zone selected.');
  } else if (button == ui.Button.CLOSE) {
  // User clicked X in the title bar.
    ui.alert('You closed the dialog.');
  };


  if (button == ui.Button.OK) {
    var url = "https://api.weather.gov/alerts?zone=AKZ"+zone; //TXZ
    //add zone here
    var slide = SlidesApp.getActivePresentation().getSelection().getCurrentPage();

    var res = UrlFetchApp.fetch(url);
    var content = res.getContentText();
    var json = JSON.parse(content);
    Logger.log(json);

    var newOrPrevious = 0;

    if (json['features'].length != 0){
      //test+=1;
      // if features exist then there is a warning for that zone:
      //var zone = json['features'][0]['properties']['geocode']['UGC'];
      var description = json['features'][newOrPrevious]['properties']['description'];
      var description = description.replaceAll('\n',' ');
      var product = description.split("*");
    };

    if (product.length > 1){
      var title = "WHAT"
      var shapes = slide.getShapes();
      var s = shapes.filter(function(e) {return e.getTitle() == title});
      if (s.length > 0) {
        s[0].getText().setText(product[1]);
      } else {
        var shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 45, 80, 280, 60).setTitle(title);
        var textRange = shape.getText();
        textRange.setText(product[1]);
        textRange.getTextStyle().setFontSize(11).setForegroundColor('#ffffff');
      }

      var title = "WHERE"
      var shapes = slide.getShapes();
      var s = shapes.filter(function(e) {return e.getTitle() == title});
      if (s.length > 0) {
        s[0].getText().setText(product[2]);
      } else {
        var shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 45, 205, 280, 60).setTitle(title);
        var textRange = shape.getText();
        textRange.setText(product[2]);
        textRange.getTextStyle().setFontSize(11).setForegroundColor('#ffffff');
      }

      var title = "WHEN"
      var shapes = slide.getShapes();
      var s = shapes.filter(function(e) {return e.getTitle() == title});
      if (s.length > 0) {
        s[0].getText().setText(product[3]);
      } else {
        var shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 45, 265, 280, 60).setTitle(title);
        var textRange = shape.getText();
        textRange.setText(product[3]);
        textRange.getTextStyle().setFontSize(11).setForegroundColor('#ffffff');
      }

      var title = "IMPACTS"
      var shapes = slide.getShapes();
      var s = shapes.filter(function(e) {return e.getTitle() == title});
      if (s.length > 0) {
        s[0].getText().setText(product[4]);
      } else {
        var shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 45, 335, 280, 60).setTitle(title);
        var textRange = shape.getText();
        textRange.setText(product[4]);
        textRange.getTextStyle().setFontSize(11).setForegroundColor('#ffffff');
      }

      if (product.length > 4){
        var title = "ADDITIONAL"
        var shapes = slide.getShapes();
        var s = shapes.filter(function(e) {return e.getTitle() == title});
        if (s.length > 0) {
          s[0].getText().setText(product[5]);
        } else {
          var shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 45, 125, 280, 60).setTitle(title);
          var textRange = shape.getText();
          textRange.setText(product[5]);
          textRange.getTextStyle().setFontSize(11).setForegroundColor('#ffffff');
        }
      };
    } else {
      SlidesApp.getUi().alert(' '+description+' View API for details here: '+url+'.'); 
    };
     
  };
};



//superHeroes['members'][1]['powers'][2]
//https://developer.mozilla.org/en-US/docs/Learn/JavaScript/Objects/JSON

//https://api.weather.gov/products/types/WSW/locations/AJK
//var date = Utilities.formatDate(new Date(), "America/Anchorage","MM/dd/yyyy h:mm a z");
//var url = "https://api.weather.gov/alerts/active?zone=AKZ019";
// https://api.weather.gov/products/
//https://developers.google.com/google-ads/scripts/docs/features/dates
//https://stackoverflow.com/questions/55750034/adding-todays-date-to-presentation-slide-added-but-cant-update
