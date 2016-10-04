<body onload="JavaScriptTable();">

<script>

function JavaScriptTable() {
    var oTable = document.createElement('<table>');
    var oTBody = document.createElement('tbody');
    var oRow = document.createElement('tr');
    var oCell = document.createElement('td');
    var oText = document.createTextNode('This is a table of one cell created by JavaScript');

    oTable.appendChild(oTBody);
    oTBody.appendChild(oRow);
    oRow.appendChild(oCell);
    oCell.appendChild(oText);
    oCell.appendChild(oText);

    oTable.style.border = '2px double';
    document.body.appendChild(oTable);
}
</script>
