/* jshint esversion: 8 */

function addElement(elt, pInst, media) {
  const node = pInst._userNode ? pInst._userNode : document.body;
  node.appendChild(elt);
  const c = media
    ? new p5.MediaElement(elt, pInst)
    : new p5.Element(elt, pInst);
  pInst._elements.push(c);
  return c;
}

p5.prototype.createFileInput2 = function (callback, multiple = false) {
  p5._validateParameters('createFileInput', arguments);

  const handleFileSelect = function (event) {
//-----------------------------修正開始
    msg = "ファイル読み込み中です";
    let fs = [];
    for (const file of event.target.files) {
        if (file.name.slice(-4) === '.csv') {
          p5.File._load(file, callback); //最初にCSVファイルを呼び出す 
        } else {
          fs.push(file);
        }
    }
    for (const f of fs) {
      p5.File._load(f, callback);
    }
//-----------------------------修正終了
  };

  // If File API's are not supported, throw Error
  if (!(window.File && window.FileReader && window.FileList && window.Blob)) {
    console.log(
      'The File APIs are not fully supported in this browser. Cannot create element.'
    );
    return;
  }

  const fileInput = document.createElement('input');
  fileInput.setAttribute('type', 'file');
  if (multiple) fileInput.setAttribute('multiple', true);
  fileInput.addEventListener('change', handleFileSelect, false);
  return addElement(fileInput, this);
};
