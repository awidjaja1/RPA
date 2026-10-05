window.AA360functions = function(arg){
    var SPAN_ID ="DERIVED_HR_GB_GB_COUNT_NBR$0";
    var iframeid = "ptifrmtgtframe";
    var iframe = document.getElementById(iframeid);
    var frameWin = iframe.contentWindow;
    var frameDoc = iframe.contentDocument || frameWin.document;

    
    function getInitialCount(){
      try{
        window._calcDone = false;
        window._countBefore = null;
        window._countAfter = null;

        console.log('Getting initial count');
        var el = frameDoc.getElementById(SPAN_ID);
        window._countBefore = el.innerText;
        return window._countBefore;
      
      }catch(e){
        return e.message;
      }
    }
    function setXHR(){
        try{
          console.log('Setting XHR');
          if(!frameWin._origPen){
            frameWin._origPen = frameWin.XMLHttpRequest.prototype.open;
          }
            frameWin.XMLHttpRequest.prototype.open = function(){
                this.addEventListener('load',function(){
                    var text = this.responseText || '';
                    if (text.indexOf('DERIVED_HR_GB_GB_COUNT_NBR$0') > -1) {
                        window._calcDone = true;
                    }
                });
                return frameWin._origPen.apply(this, arguments);
            };
            return "Installed simple XHR monitor";
        }catch(e){
            return e.message;
        }

    }
    function getCalcDone(){
        console.log('Getting window._calcDone: '+window._calcDone);
        return window._calcDone; 
    }
    function checkDone(){
        try{
            console.log('Checking window._calcDone: '+window._calcDone);
            if (window._calcDone !== true) { 
                return '__waiting__'; 
            }
            else{
                var el2 = frameDoc.getElementById('DERIVED_HR_GB_GB_COUNT_NBR$0');
                window._countAfter = el2 ? el2.innerText : '';
                console.log("val: "+ window._countAfter );
                return window._countAfter;
            }
        }catch(e){
            return e.message;
        }
    }
  function cleanUp(){
    try{

      window._calcDone = false;
      window._countBefore = null;
      window._countAfter = null;
      if(frameWin._origPen){
        frameWin.XMLHttpRequest.prototype.open = frameWin._origPen;
        frameWin._origPen = null;
      }
      return 'XHR monitor cleaned up';
    }catch(e){return e.message;}
  }



  if(arg ==="getInitialCount"){
    return getInitialCount();
  }
  if(arg === "setXHR"){
    return setXHR();
  }
  if(arg === "getCalcDone"){
    return getCalcDone();
  }
  if(arg === "checkDone"){
    return checkDone();
  }
  if(arg ==="cleanUp"){
    return cleanUp();
  }


};