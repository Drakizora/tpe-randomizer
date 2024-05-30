window.onload = getCat;

function getCat() {

  var url = "https://docs.google.com/spreadsheets/d/1_cd5S8q-biIABfCYeQf1X_I1XjdIKglEeNsV7Z6yX2c/export?format=xlsx&gid=0";

  fetch(url)
    .then(response => {
      if (!response.ok) {
        throw new Error('Something went wrong!');
      }
      return response.arrayBuffer();
    })
    .then(data => {
      var workbook = XLSX.read(data, { type: 'array' });
      var firstSheetName = workbook.SheetNames[0];
      var worksheet = workbook.Sheets[firstSheetName];
      var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      jsonData = jsonData.map((row, rowIndex) => {
        return row.map((cell, colIndex) => {
          let cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
          let cellObject = worksheet[cellAddress];
          if (cellObject && cellObject.l) {
            return cellObject.l.Target;
          } else {
            return cell;
          }
        });
      });

      // Randomly selecting character
      var targetRowIndex = jsonData.findIndex(row =>
        row.some(cell => typeof cell === 'string' && cell.includes("Vis."))
      );
      var firstIndex = targetRowIndex + 1;

      var targetRowIndex = jsonData.findIndex(row =>
        row.some(cell => typeof cell === 'string' && cell.includes("DECEASED CHARACTERS"))
      );
      var lastIndex = targetRowIndex - 2;

      var index = getRandomIndex(firstIndex, lastIndex);
      var selection = jsonData[index];

      // Retrieving information
      var cat = {
        Name: selection[3],
        Application: getImageLink(selection[1]),
        Biography: selection[2],
        Player: selection[selection.length - 1]
      }

      document.getElementById('name').textContent = cat.Name;
      document.getElementById('application').setAttribute("src", cat.Application);
      document.getElementById('biography').setAttribute("href", cat.Biography);
      document.getElementById('player').textContent = cat.Player;
    })
    .catch(error => {
      console.error(error);
    });
}

function getRandomIndex(lower, upper) {
  return Math.floor(Math.random() * (upper - lower + 1)) + lower;
}

function getImageLink(url) {
  var startIndex = url.indexOf('/d/') + 3;
  var endIndex = url.indexOf('/view');
  if (startIndex !== -1 && endIndex !== -1) {
    var id = url.substring(startIndex, endIndex);
  }
  else {
    return null;
  }
  imageLink = "https://drive.google.com/thumbnail?id=" + id + "&sz=w1000";
  return imageLink;
}
