function getFileName(id,label,text) {
  var afilepath = document.getElementById(id).value;
  var splitpath = afilepath.split("\\")

  document.getElementById(label).textContent = text + splitpath[splitpath.length - 1];
}

document.getElementById('src_tgt').addEventListener('change', getFileName.bind(null, 'src_tgt', 'src_tgt_text',
'File with source-target mappings: '), false);
document.getElementById('process_file').addEventListener('change', getFileName.bind(null, 'process_file', 'process_file_text',
'File to be processed: '), false);
