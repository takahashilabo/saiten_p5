/* jshint esversion: 8 */

let w, h;
let mag_disp;
const mag_excel = 0.25; //Excelに解答画像を貼り付けるときの縮小率
let dragged = false;
let vec;
let arr = [];
let pre_key = "";
let ans_no = 1;
let input;
let img = null;
let mode = 0; //0:初期, 1:キリトリ位置指定, 2:キリトリ実行
let wb;
let ws;
const CSV_FILE = 'trimData.csv'
let uploadfile_count = 0;
let uploadfile_names = [];
let uploadfile = {};
let msg = "";

function setup() {
  createCanvas(640, 320);
  init_excel();
}

function init_excel() {
  wb = new ExcelJS.Workbook(); 
  ws = wb.addWorksheet('a');
}

function draw() {
  if (mode == 0) {
    background(255);
    fill(0); textSize(20);
    text('1か2を押してください（1:キリトリ位置設定, 2:キリトリ実行)', 0, 20);
  } else if (mode == 1) { //キリトリ位置設定
    if (img) {
      image(img, 0, 0, width, height);
      strokeWeight(2);
      for (let a of arr) {
        stroke(0, 0, 255); noFill();
        rect(a.start_x, a.start_y, a.end_x - a.start_x, a.end_y - a.start_y);
        noStroke(); fill(0, 0, 255);
        text(a.ans_no, a.start_x + 5, a.start_y + 20);
      }
    
      if (dragged) {
        noFill(); stroke(0, 0, 255); strokeWeight(2);
        rect(vec.x, vec.y, mouseX - vec.x, mouseY - vec.y);
      }
    }
  } else if (mode == 2) { //キリトリ実行
    if (uploadfile_count > 0) {
      msg = (uploadfile_count < input.elt.files.length - 1) ? `解答欄切り取り中（${uploadfile_count}/${input.elt.files.length - 1}）` : '解答欄切り取り完了！'; 
    }
    background(255);
    fill(0); textSize(20);
    text(msg, 30, 50);
  }
}

function mousePressed() {
  if (mode == 1) {
    if (img) {
      dragged = true;
      vec = createVector(mouseX, mouseY);
    }
  }
}

function mouseReleased() {
  if (mode == 1) {
    if (img) {
      dragged = false; 
      arr.push({'start_x':vec.x, 'start_y':vec.y, 'end_x':mouseX, 'end_y':mouseY, 'ans_no':ans_no++});
    }
  }
}

function keyPressed() {
  if (mode == 0) {
    if (key == '1') {
      mode = int(key);
      background(255);
      input = createFileInput(handleFile_mode1);
      input.position(0, 0);
    }
    if (key == '2') {
      mode = int(key);
      background(255);
      input = createFileInput2(handleFile_mode2, true); //複数ファイル選択可（CSVファイル＋解答画像群）
      input.position(0, 0);
    }
  } else if (mode == 1) {
    if (img) {
      if (keyCode == ESCAPE) {
        arr.pop();
        if (ans_no > 1) { ans_no--; }
        return;
      }
      
      if (key == 's' || key == 'S') {
        saveCSV();
      }
      
      if (key >= '1' && key <= '9') {
        if (pre_key == "") {
          pre_key = key;
        } else {
          let row = int(key);
          let col = int(pre_key);
          pre_key = "";
          let a = arr.pop();
          if (ans_no > 1) { ans_no--; }
          let dx = (a.end_x - a.start_x) / col;
          let dy = (a.end_y - a.start_y) / row;
          for (let i = 0; i < row; i++) {
            for (let j = 0; j < col; j++) {
              arr.push({
                'start_x':a.start_x + j * dx,
                'start_y':a.start_y + i * dy,
                'end_x':a.start_x + (j+1) * dx,
                'end_y':a.start_y + (i+1) * dy,
                'ans_no':ans_no++});
            }
          }
        }
      }
    }
  }
}

function handleFile_mode1(file) {
  if (file.type === 'image') {
    img = loadImage(file.data, e => {
      mag_disp = displayWidth / e.width;
      w = e.width * mag_disp;
      h = e.height * mag_disp;
      resizeCanvas(w, h);
    });
  } else {
    img = null;
  }
  input.remove();
}

function handleFile_mode2(file) {
  if (file.name.slice(-4) === '.csv') { //必ず最初にCSVファイルがくる
    csv_to_arr(file.data);
    for (let f of input.elt.files) {
      uploadfile_names.push(f.name); //アップロード予定のファイル一覧
    }
    uploadfile_names.sort(); //Excelシート上に学籍番号順に並べるため辞書順にソートする
    for (let i = 0; i < uploadfile_names.length; i++) {
      uploadfile[uploadfile_names[i]] = i; //ファイル名ごとのExcel行位置をセット
    }
  }
  else if (file.type === 'image') { //次に画像ファイルがくる
    loadImage(file.data, e => {
      attach_ans_to_excel(file.name, e); //e = PImage
      if (++uploadfile_count >= input.elt.files.length - 1) { //-1 : CSVファイルを除く意味
        save_xlsx();
        init_excel(); //free
      }
    });
  }
  input.remove();
}

function attach_ans_to_excel(filename, e) {
  e.loadPixels();
  let col = 1;
  let col_width = [];
  let max_row_height = 0;
  let row = uploadfile[filename]; //filenameのデータをセットするExcelシート行位置
  for (let a of arr) {
    let p = e.get(a.start_x, a.start_y, a.end_x - a.start_x, a.end_y - a.start_y);
    //p.filter(ERODE); //文字見やすくするため
    p.resize(p.width * mag_excel, 0);
    let logo = wb.addImage({base64: p.canvas.toDataURL(), extension: 'jpg'});
    ws.addImage(logo, {
      tl: { col: col - 1, row: row },
      ext: { width: p.width, height: p.height },
    });
    ws.getRow(row + 1).getCell(col).fill =  { type: 'pattern', pattern: 'solid', fgColor: { argb:'FF111111' }};
    col_width.push(p.width);
    col_width.push(40);
    let b = ws.getRow(row + 1).getCell(col + 1);
    b.value = 1; //1点
    b.alignment = { vertical: 'top', horizontal: 'left' };
    b.font = {name: 'ＭＳ Ｐゴシック', color: { argb: 'FFFFFFFF' }, family: 1, size: 20};
    b.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF111111' }};
    if (max_row_height < p.height * 0.6) {
      max_row_height = p.height * 0.6;
    }
    col += 2;
  }
  
  ws.columns.forEach(function (column) {
    column.width = col_width.shift() * 0.14;
  });

  ws.getRow(row + 1).height = max_row_height;
}

function csv_to_arr(data) {
  arr = [];
  lines = data.split("\n");
  let keys = lines[0].split(",");
  for (let i = 1; i < lines.length; i++) {
    let l = lines[i].split(",");
    if (l.length == 1) continue; //最終ゴミ行削除のため
    h = {};
    for (let j = 0; j < keys.length; j++) {
      h[keys[j]] = (j > 0) ? int(l[j]) : l[j]; //0:name, 1>座標
    }
    arr.push(h);
  }
}

function saveCSV() {
  let table = new p5.Table();
  for (let a of ["tag", "start_x", "start_y", "end_x", "end_y"]) {
    table.addColumn(a);
  }
  let row = 0, col = 0, q_no = 0;
  for (const a of arr) {
    table.addRow();
    col = 0;
    let tag = (q_no == 0) ? "name" : 'Q_' + ('0000' + q_no).slice(-4);
    q_no++;
    table.set(row, col++, tag);
    table.set(row, col++, int(a.start_x / mag_disp));
    table.set(row, col++, int(a.start_y / mag_disp));
    table.set(row, col++, int(a.end_x / mag_disp));
    table.set(row++, col++, int(a.end_y / mag_disp));
  }
  save(table, CSV_FILE);
}

async function save_xlsx() { //Excelファイル生成＆ダウンロード
  const uint8Array = await wb.xlsx.writeBuffer();
  const blob = new Blob([uint8Array], { type: 'application/octet-binary' });
  const a = document.createElement('a');
  a.href = (window.URL || window.webkitURL).createObjectURL(blob);
  a.download = 'output.xlsx';
  a.click();
  a.remove();
}
