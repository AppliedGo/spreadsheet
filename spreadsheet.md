# Applied Go Quick Bits: 


<!-- 
course

In this lecture, we will look at accessing spreadsheet data, using only functions from the standard library.
-->


Spreadsheet data is everywhere. You can find it in Excel sheets as well as when downloading business data from a website. 

Package [`encoding/csv`](https://golang.org/pkg/encoding/csv/) from the Go standard library can help you processing that data and produce statistics, reports or other kinds of output from it. Here is how.

Let's assume we work at a stationery distributor. Every evening, we receive a spreadsheet containing the orders of the day for review. The data looks like this:

Date       | Order ID | Order Item   | Unit Price | Quantity
-----------|----------|--------------|------------|----------
2017-11-17 | 1        | Ball Pen     | 1.99       | 50
2017-11-17 | 2        | Notebook     | 12.99      | 10
2017-11-17 | 3        | Binder       | 4.99       | 25
2017-11-20 | 4        | Pencil       | 0.99       | 100
2017-11-20 | 5        | Sketch Block | 2.99       | 40
2017-11-22 | 6        | Ball Pen     | 1.99       | 30
2017-11-23 | 7        | Sketch Block | 2.99       | 20
2017-11-24 | 8        | Ball Pen     | 1.99       | 60

We're interested in some information that is not directly contained in the data; especially, we want to know 

* the total price for each order,
* the total sales volume, 
* and the number of ball pens sold.

As we get a new copy every day, creating formulas within the spreadsheet is not an option. Instead, we decide to write a small tool to do the calculations for us. Also, the tool shall add the result to the table and write a new spreadsheet file.

Let's dive into the Go code.

We start by declaring the `main` package and a `main` function. In `main()`, we sketch out our program flow:

* Read the CSV file,
* calculate the desired numbers, and
* write the results to a new CSV file.

<!--
course

The "colon-equal sign" is a "create-and-assign" operator. It puts the variable `rows` into existence and assigns the result of the call to `readOrders()` to it. The variable's type is inferred from the function's result type.
-->

Our first step is to export the spreadsheet data to CSV. To make things a bit more complicated, we export the data with a semicolon as the column separator.

Our data now looks like this:

```
Date;Order ID;Order Item;Unit Price;Quantity
2017-11-17;1;Ball Pen;1.99;50
2017-11-17;2;Notebook;12.99;10
2017-11-17;3;Binder;4.99;25
2017-11-18;4;Pencil;0.99;100
2017-11-18;5;Sketch Block;2.99;40
2017-11-19;6;Ball Pen;1.99;30
2017-11-19;7;Sketch Block;2.99;20
2017-11-19;8;Ball Pen;1.99;60
```

We can see a header row and data rows, with data separated by semicolons. 

To read this data, we first open the file for reading, using `os.Open()`. 

If the file cannot be opened, `os.Open()` returns an error that we need to handle.

<!---
course

Go's error handling philosophy is to check and handle errors immediately when they occur. At first, this might just seem to produce code that is overly verbose. However, this approach has a huge advantage: It forces the developer to think about every error case in their code. As a result, Go code tends to be very robust, and bugs can be located much easier.
-->

Usually we would return the error to the caller and handle all errors in function `main()`. However, this is just a small command-line tool, and so we use `log.Fatal()` instead, in order to write the error message to the terminal and exit immediately.


After this point, the file has been successfully opened, and we want to ensure that it gets closed when no longer needed, so we add a deferred call to `f.Close()`.

<!--
course

The `defer` directive defers a function call until the point where the current function exits. This has a very welcome effect: No matter where, when, and why the function exits, the deferred call is always executed. The `readOrders()` function has three exit points, and without a `defer` directive, we would have to ensure to call `f.Close()` on each of these exit points.
-->

To read in the CSV data, we create a new CSV reader that reads from the input file.


<!--
course

A "Reader" is a central concept in Go. It can abstract away the input source, so that the same code can be used for reading from a file, a network connection, a string, or any other character stream.
-->

The CSV reader is aware of the CSV data format. It separates the input stream into rows and columns, and returns a slice of slices of strings. 
We can even adjust the reader to recognize a semicolon, rather than a comma, as the column separator.

<!--
course

A "slice" is about the same as a dynamic array in other languages. It has no fixed size but rather grows as needed when data is appended. 
-->

Again, we check for any error, and finally we can return the rows.

To process the data, we loop over the rows, and read from or write to each row slice as needed.

The `for...range` loop iterates over the outer slice and yields the index of each element of the slice.

<!--
course

A range loop usually yields two values: the loop index, and a copy of the element at that index. However, we only need the index value here, and so we can ignore the optional second value.
-->

The first row is the header row. Here, we only want to add a new header for the column that holds the total prices.

<!--
course

`append` is a built-in function that adds a new element to a slice, and expands the capacity of the slice if required. 
-->

From the next row onwards, we calculate the total price, sum up all prices, and count the number of ball pens being ordered.

This is fairly straightforward, as we know the indexes of the item name, the unit price, and the quantity. The only difficulty is that all columns are string values but we need the price and quantity values as numeric values. 

<!--
course

Go is a statically typed language and does not allow assigning a value of a given type to a variable of another type without explicit type conversion. Static typing helps catching bugs at compile time, so that they cannot turn into runtime errors that are usually hard to track down.
-->

Another obstacle we are facing here is that the prices are floating-point values but for financial calculations, we want to use precise integer calculation only. Luckily, the [`strings`](https://golang.org/pkg/strings) and [`fmt`](https://golang.org/pkg/fmt/) packages have got us covered.

First, we remove the decimal point from the price using `strings.Replace()`, and then we can convert the strings to integers using `strconv.Atoi`.

As these operations can fail, we have to ensure to handle all errors.

After doing our calculations, we need to turn the results back into a floating-point value represented as a string. For this we write a little helper function that calculates the integer part and the fractional part from the result, and writes both values, separated by a decimal point, into a string, using `fmt.Sprintf()`.

Now we can append the total price to each row, and write the overall sum of all orders into a new row, as well as the number of ball pens being ordered.

Finally, we write the result to a new file, using `os.Create()`... and a CSV writer that knows how to turn the slice of slices of strings back into a proper CSV file. Note that we do not set the separator to semicolon here, as we rather want to get a standard CSV format this time.


When running this code, the output file should look like this:

```
Date,Order ID,Order Item,Unit Price,Quantity,Total
2017-11-17,1,Ball Pen,1.99,50,99.50
2017-11-17,2,Notebook,12.99,10,129.90
2017-11-17,3,Binder,4.99,25,124.75
2017-11-18,4,Pencil,0.99,100,99.0
2017-11-18,5,Sketch Block,2.99,40,119.60
2017-11-19,6,Ball Pen,1.99,30,59.70
2017-11-19,7,Sketch Block,2.99,20,59.80
2017-11-19,8,Ball Pen,1.99,60,119.40
,,,Sum,,811.65
,,,Ball Pens,140,
```

And we can open it in our spreadsheet app, or in a CSV viewer, to get a nicely formatted table. 

Here we can see our new Totals column, as well as the two new rows that show the overall sum and the number of ball pens ordered.

<!--
course

Finally, let's deploy our tool. The Go compiler makes this easy as pie. Just type

    go build

inside the directory where the source code is, and you will get a binary in the same directory, ready to copy to wherever you want.

Or type 

    go install

to create the binary in `$GOPATH/bin`. If you added `GOPATH` to your `PATH` variable, you can run the tool from any directory.

-->

You can do a lot more with the spreadsheet data. Try, for example, handling orders with multiple items, or add columns like discount, base price, or shipping cost per order, and calculate new total prices as well as your revenue from that.

Happy coding! `ʕ◔ϖ◔ʔ` 
