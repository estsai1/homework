// from data.js
var tableData = data;

// YOUR CODE HERE!

// Build complete table at the beginning
function init() {

    // Get reference to table body
    var tbody = d3.select("tbody");

    // Console.log the data
    console.log(tableData)

    // Append table to the webpage
    // I borrowed the code from 03-Evr_D3_Table because it's a very elegant solution, I hope that's acceptable
    tableData.forEach((report) => {
        var row = tbody.append("tr");
        Object.entries(report).forEach(([key, value]) => {
            var cell = row.append("td");
            cell.text(value);
        });
    });

}

// Select the button and form
var button = d3.select("#button");
var form = d3.select("#form")

// Create event handlers
button.on("click", runEnter);
form.on("submit", runEnter);

// Function to run for both events
function runEnter() {

    d3.event.preventDefault();

    // Clear the table
    // Found this method from https://stackoverflow.com/questions/7271490/delete-all-rows-in-an-html-table
    var table = document.getElementById("tablebody");
    table.innerHTML = "";

    // Select input element and get raw HTML node
    var inputElement = d3.select("#datetime");

    // Get value property of the input element
    var inputValue = inputElement.property("value");

    console.log(inputValue);

    var filteredData = tableData.filter(entry => entry.datetime == inputValue);

    console.log(filteredData);

    // Rebuild the table with filtered data
    var tbody = d3.select("tbody");

    filteredData.forEach((report) => {
        var row = tbody.append("tr");
        Object.entries(report).forEach(([key, value]) => {
            var cell = row.append("td");
            cell.text(value);
        });
    });

}

init();