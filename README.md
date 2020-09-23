<div align="center">

## Using Collections in Visual Basic \#2 \- Jet Database Collections


</div>

### Description

This article attempts to explain the Microsoft Jet collections and how you can use them in really useful ways. If you don't know about Jet collections, this is well worth reading.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-09-14 05:59:58
**By**             |[Matthew Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-roberts.md)
**Level**          |Intermediate
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD99639182000\.zip](https://github.com/Planet-Source-Code/matthew-roberts-using-collections-in-visual-basic-2-jet-database-collections__1-11529/archive/master.zip)





### Source Code

<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<META NAME="Generator" CONTENT="Microsoft Word 97">
<TITLE>Jet Collections</TITLE>
<META NAME="Template" CONTENT="D:\OFFICE97\OFFICE\html.dot">
</HEAD>
<BODY LINK="#0000ff" VLINK="#800080">
<B><FONT FACE="Arial" SIZE=5><P ALIGN="CENTER">Using Collections in Visual Basic </P>
</B></FONT><I><FONT FACE="Arial" SIZE=2><P ALIGN="CENTER">Part 2 - Jet Database Collections</P>
<P ALIGN="CENTER"> </P>
</I><P>Most Visual Basic developers are familiar with the Jet Database Engine. While it receives a lot of flack from developers who work with more powerful systems such as Oracle or SQL Server, Jet has a lot of really good features that make it ideal for a desktop application. Besides, we VB Developers are used to sneers and comments from "hard core" language programmers. "What is that? A string parameter? Why, in C++, we don't pass strings! We pass pointers to memory addresses that contain null terminated string arrays! That's how real men handle strings!"</P>
<P>Yea, whatever.</P>
<P>Obviously we VB developers aren't interested in doing things the hard way, and Jet is a wonderful way to avoid it while still having a high level of control over your data. Before I go into the wonderful benefits of the Jet Database Engine, I think it is appropriate to point out that if you are in the habit of using data controls on your forms to access databases, you are greatly limiting your freedom to work with data, and you are bypassing many of the most useful things about Jet. At the risk of sounding like the C++ developer I was just making fun of, you really should take the time to learn DAO or ADO. If there is any interest in a general "This is how you use DAO/ADO" tutorial out there, let me know and I will work one up.</P>
<P>This tutorial won't go over how to use DAO or ADO for data access. Since DAO seems to be the most common method for accessing Jet data right now, I will give all of my code examples in DAO. If you don't know how to use DAO yet, maybe this article will convince you that it is worth learning. There have been entire books written on using the Jet Database Engine, so to try to cover the "how to" basics AND Collections here would get pretty long winded. So I am going to stick to collections.</P>
<P>Now, on with the tutorial:</P>
<P>Microsoft Jet is not actually a thing. It is more like a format. A Jet database consists of a single file with many internal elements. You Access Developers out there will be familiar with the concept of a single Access .mdb file containing many different objects. While it is useful in Access to have forms, macros, and reports in a single file, it is kind of pointless with Visual Basic. You have no way to use those objects from the VB environment, so they are just filler. Therefore, we are going to focus on the Table and Query objects. Don't get too hung up on how Jet stores all of these things in a single file and keeps up with it all, just trust that it does and go with it. For the technically curious, .mdb files are similar to a miniature file system within a single file "wrapper". </P>
<P>So, on to the meat of this thing. Jet Database Collections. </P>
<P>If you read my other tutorial on </FONT><A HREF="http://www.planetsourcecode.com/xq/ASP/txtCodeId.9349/lngWId.1/qx/vb/scripts/ShowCode.htm"><FONT SIZE=2>collections</FONT></A><FONT FACE="Arial" SIZE=2>, or if you have worked with collections before, this will not seem totally new to you. If not, you can probably hang in there anyway. These examples aren't tough.</P>
<P>Microsoft Jet is, as mentioned, a collection of database objects in a single file. These objects have a <B>hierarchy</B>. This just means that there are top level and lower level members, and the top level ones "contain" lower level ones. In Jet, the highest level object is the Database object. It is, for all practical purposes, the file itself. Think of it as a big box. Within that box we see other objects. The ones we are concerned with are Tables and Queries. </P>
<P>When Jet stores a table or query, it actually stores a set of information that acts as a table definition. It describes the table to the Jet database engine, and the Jet engine creates it when it needs it. Think of it as a template. The names Microsoft chose to give these objects are a little puzzling unless you know that they are definitions. They are called <B>TableDefs</B> and <B>QueryDefs</B>. They are essentially identical from the collections point of view, so we will concentrate on tablesdefs and then wrap it up with querydefs.</P>
<P>So enough of this technical stuff, how about some code. </P>
</FONT><FONT FACE="Arial" SIZE=1><P>(NOTE: DAO requires a reference to the Data Access Object in the References of your project. For information on how to add a reference to DAO, see VB Help and search for "</FONT><FONT SIZE=2>Creating a Reference to an Object</FONT><FONT FACE="Arial" SIZE=1>))</P>
</FONT><FONT FACE="Arial" SIZE=2><P> </P>
<P>Take this example:</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2>ShowCustomers()</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	Dim </FONT><FONT FACE="Courier New" SIZE=2>dbCustomers</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> As </FONT><FONT FACE="Courier New" SIZE=2>Database</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	Dim </FONT><FONT FACE="Courier New" SIZE=2>rsCustomers</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> As </FONT><FONT FACE="Courier New" SIZE=2>Recordset</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#008000"><P>	'	Create your data objects and open the table "Customers"</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	Set </FONT><FONT FACE="Courier New" SIZE=2>dbCustomers</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> </FONT><FONT FACE="Courier New" SIZE=2>= OpenDatabase ("C:\Program Files\CustomerInfo\Customers.mdb")</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	</FONT><FONT FACE="Courier New" SIZE=2>Set rsCustomers = dbCustomers.OpenRecordset ("Customers")</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#008000">'	List some information from the database to the debug window</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	Debug.Print </FONT><FONT FACE="Courier New" SIZE=2>rsCustomers!LastName</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	Debug.Print </FONT><FONT FACE="Courier New" SIZE=2>rsCustomers!FirstName</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	Debug.Print </FONT><FONT FACE="Courier New" SIZE=2>rsCustomers!PhoneNumber</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P> </P>
<P>	</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#008000">'	Always clean up after you are done!</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	set </FONT><FONT FACE="Courier New" SIZE=2>dbCustomers</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> = Nothing</P>
<P>	set </FONT><FONT FACE="Courier New" SIZE=2>rsCustomers</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> = Nothing </P>
<P>End Sub</P>
</FONT><FONT FACE="Arial" SIZE=2><P>This sub simply opens a Jet database and displays three values from it. This is pretty easy and straightforward. But there is a problem with this type of code. You have to have prior knowledge of what is in the data file. You have to know the table name, and within the table, you have to know the field names. You may even need to know if those fields are number or string fields. Is PhoneNumber a string or Long datatype? How can you tell? Do we even care?</P>
<P>Answer: Probably not. Most of the time that we access databases, we already know the field names and datatypes. So what is my point? My point is, you may not know. I recently created a small project that would allow you to select a table from an Access .mdb file and view all of the data and in a grid. There is no possible way I could know what any random database file is going to contain. There could be any number of tables with any names, and each of those tables could have any arrangement of fields. Obviously there is a way to get to that sort of information in code without knowing it in advance. Either that or my project was a miserable failure, and I can tell you it wasn't...just a modest one. There is a way to examine any Jet database and determine its elements. This method is collections (FINALLY!).</P>
<P>If you remember, I said earlier that the highest level object in a Jet database is the Database object. That means that if we want to refer to anything within the database, you must reference it THROUGH this object. But how do you do THAT? We already have. Look at the code above and you will see this line:<BR>
<BR>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080">	Set </FONT><FONT FACE="Courier New" SIZE=2>rsCustomers</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> = </FONT><FONT FACE="Courier New" SIZE=2>dbCustomers.OpenRecordset ("Customers")</P>
</FONT><FONT FACE="Arial" SIZE=2><P>This line tells VB to create a new Recordset object based on dbCustomers, using the table Customers. To refer directly to that table in code, you could use this syntax:</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	dbCustomers.TableDefs("Customers")</P>
</FONT><FONT FACE="Arial" SIZE=2><P>Because what you are really telling it to do is to look at the table named "Customers", which is part of the TableDefs COLLECTION in the database dbCustomers. The TableDefs collection contains all of the tables in the database...even the super-secret hidden ones that Jet uses internally to manage the data. Hidden tables will begin with mSys. You will see them later on.</P>
<P>But wait! In the example above, I just did it the hard way. I still had to know the name of the table...or did I? Although you can refer to the tables in the manner that I did, you don't have to. All collections in Visual Basic are enumerated. That means that they are basically a glorified array. And as you know, you can refer to the elements of an array with an index number. For example, to find out what the 5the element in a string array is, you could do this:</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>strTest = strTestArray(4)</P>
</FONT><FONT FACE="Arial" SIZE=2><P>(Remember, arrays are zero-based unless you specify the Option Base explicitly...just a reminder).</P>
<P>So to get the first element in the array, you could say this:</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	strTest = strTestArray(0)</P>
</FONT><FONT FACE="Arial" SIZE=2><P>Easy, right? Well then you have got the concept. You can reference tables in a Jet database the same way:</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	strTableName = dbCustomers!TableDefs(0).Name</P>
<P>	Debug.Print strTableName</P>
</FONT><FONT FACE="Arial" SIZE=2><P>		</P>
<P>This will return the name of the table - Customers.</P>
<P>Wait! This is getting cool! That means that if you know the index number, you can get the name! But how can I know the index of a particular table? You can't. But as you will see, it doesn't matter, because you can use the For...Next command to go through them all.</P>
<P>Check this out:</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2>ListTables()</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	</P>
<P>	Dim </FONT><FONT FACE="Courier New" SIZE=2>dbTableList</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> As </FONT><FONT FACE="Courier New" SIZE=2>Database</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	Dim </FONT><FONT FACE="Courier New" SIZE=2>intTableNumber</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> As </FONT><FONT FACE="Courier New" SIZE=2>Integer</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	Dim </FONT><FONT FACE="Courier New" SIZE=2>strTableName</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> as </FONT><FONT FACE="Courier New" SIZE=2>String</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	Set </FONT><FONT FACE="Courier New" SIZE=2>dbTableList = OpenDatabase ("C:\program files\customerinfo\ customers.mdb")</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	For </FONT><FONT FACE="Courier New" SIZE=2>intTableNumber = 0 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080">To </FONT><FONT FACE="Courier New" SIZE=2>dbTableList.TableDefs.Count - 1</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> </P>
<P>		</FONT><FONT FACE="Courier New" SIZE=2>strTableName = dbTableList.TableDefs(intTableNumber).Name</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>		Debug.Print </FONT><FONT FACE="Courier New" SIZE=2>strTableName</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	Next </FONT><FONT FACE="Courier New" SIZE=2>intTableNumber</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>End Sub</P>
</FONT><FONT FACE="Arial" SIZE=2><P>There are a couple of things to note here. The first is that I used 0 to dbTableList.TableDefs.Count </FONT><B><FONT FACE="Arial" SIZE=2 COLOR="#ff0000">-1</B></FONT><FONT FACE="Arial" SIZE=2>. All collections have a built-in property "Count" which contains the number of elements in the collection. This is just like the Recordset's RecordCount property. If you have ever done this:</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	intRecords = rsCustomers.RecordCount</P>
</FONT><FONT FACE="Arial" SIZE=2><P>Then you have used the Count property. It always returns a number equal to the number of elements. If there are no elements, it will return 0. </P>
<P>The next thing to note is the use of the Name property. As with all object in VB, each element in the TableDefs collection can have an associated Name. This is exactly what you were referring to earlier when you said:</P>
</FONT><FONT FACE="Arial" SIZE=2 COLOR="#000080"><P> </P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	dbCustomers.TableDefs("Customers")</P>
</FONT><FONT FACE="Arial" SIZE=2><P>So now you see how you can get the names of tables without any prior knowledge of the database. You can also get other properties from them such as RecordCount. Take some time to explore all of the available properties...you may be surprised.</P>
<P>We now have a big part of the problem whipped. We can go into a table and list the table names in code. Cool. But what about fields? Trust me, it is EXACTLY the same concept.</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>		For </FONT><FONT FACE="Courier New" SIZE=2>intFieldNumber = 0 to rsCustomers.Fields.Count - 1</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> </P>
<P>	</FONT><FONT FACE="Courier New" SIZE=2>		strFieldName = rsFieldList.Fields.Name</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>			Debug.Print </FONT><FONT FACE="Courier New" SIZE=2>strFieldName</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>		Next intFieldNumber</P>
</FONT><FONT FACE="Arial" SIZE=2><P>This works because the TableDef object contains a Fields collection. You could combine the two examples and get a list of EVERY FIELD in EVERY TABLE in your database. Try it:</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>Private Sub </FONT><FONT FACE="Courier New" SIZE=2>ListAllTables</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080">()</P>
<P>	</P>
<P>	Dim </FONT><FONT FACE="Courier New" SIZE=2>dbTableList</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> As </FONT><FONT FACE="Courier New" SIZE=2>Database</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	Dim </FONT><FONT FACE="Courier New" SIZE=2>strTableName</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> as </FONT><FONT FACE="Courier New" SIZE=2>String</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	Dim </FONT><FONT FACE="Courier New" SIZE=2>intTableNumber</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> As </FONT><FONT FACE="Courier New" SIZE=2>Integer</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	Dim </FONT><FONT FACE="Courier New" SIZE=2>strFieldName</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> As </FONT><FONT FACE="Courier New" SIZE=2>String</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	Dim </FONT><FONT FACE="Courier New" SIZE=2>intFieldNumber</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> As </FONT><FONT FACE="Courier New" SIZE=2>Integer</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	Set </FONT><FONT FACE="Courier New" SIZE=2>dbTableList = OpenDatabase ("C:\program files\customerinfo\customers.mdb")</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	For </FONT><FONT FACE="Courier New" SIZE=2>intTableNumber = 0</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> To </FONT><FONT FACE="Courier New" SIZE=2>dbTableList.TableDefs.Count - 1</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> </P>
<P>		</FONT><FONT FACE="Courier New" SIZE=2>strTableName = dbTableList.TableDefs.Name</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>		Debug.Print </FONT><FONT FACE="Courier New" SIZE=2>strTableName</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080">			</P>
<P>		For </FONT><FONT FACE="Courier New" SIZE=2>intFieldNumber = 0 </FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080">To</FONT><FONT FACE="Courier New" SIZE=2> rsCustomers.Fields.Count - 1</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"> </P>
<P>			</FONT><FONT FACE="Courier New" SIZE=2>strFieldName = rsFieldList.Fields.Name</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>			Debug.Print </FONT><FONT FACE="Courier New" SIZE=2>strFieldName</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>		Next </FONT><FONT FACE="Courier New" SIZE=2>intFieldNumber</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	</FONT><FONT FACE="Courier New" SIZE=2>	intFieldNumber = 0 </P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>	Next </FONT><FONT FACE="Courier New" SIZE=2>intTableNumber</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>End Sub</P>
</FONT><FONT FACE="Arial" SIZE=2><P>How about that! You can make it look a little neater by adding an indention for the fields. Just change this line: </P>
</FONT><FONT FACE="Courier New" SIZE=2><P>	</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080">Debug.Print </FONT><FONT FACE="Courier New" SIZE=2>strFieldName</P>
</FONT><FONT FACE="Arial" SIZE=2><P>to</P>
</FONT><FONT FACE="Courier New" SIZE=2><P>	</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080">Debug.Print </FONT><FONT FACE="Courier New" SIZE=2>& " " & strFieldName</P>
</FONT><FONT FACE="Arial" SIZE=2><P>Your debug window will contain something like this:</P>
<P>Customers</P>
<P>	LastName</P>
<P>	FirstName</P>
<P>	PhoneNumber</P>
<P>Orders</P>
<P>	OrderNumber</P>
<P>	Amount</P>
<P>	Date</P>
<P>....</P>
<P> </P>
<P>I could go on with examples, but I bet you get the idea now. I will get your curiosity up by telling you that the Field object also contains a Properties collection. It has such information as Data Type, Length, Name, etc. That is how you were able to get the name of the field. You can access this collection like this:</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>		</FONT><FONT FACE="Courier New" SIZE=2>strFieldName = rsFieldList.Fields.Properties(2).Name</P>
</FONT><FONT FACE="Courier New" SIZE=2 COLOR="#000080"><P>		</FONT><FONT FACE="Courier New" SIZE=2>strFieldName = rsFieldList.Fields.Properties(2).Length</P>
</FONT><FONT FACE="Arial" SIZE=2><P>With that I will turn you loose to go experiment on your own. I am including the application that I wrote. It is VERY well commented, so maybe you can see how all of this database collections stuff is put to work. </P>
<P>By the way, I mentioned QueryDefs as well as TableDefs. Basically, the only difference is that you reference a saved query by referencing the QueryDefs collection instead of the TableDefs collection. Example: </P><DIR>
<DIR>
</FONT><FONT FACE="Courier New" SIZE=2><P>Set rsCustomers = dbCustomers.QueryDefs("Customers") </P></DIR>
</DIR>
</FONT><FONT FACE="Arial" SIZE=2><P>Or</P><DIR>
<DIR>
</FONT><FONT FACE="Courier New" SIZE=2><P>Set rsCustomers = dbCustomers.QueryDefs(2)</P>
</FONT><FONT FACE="Arial" SIZE=2><P> </P>
<P> </P></DIR>
</DIR>
<P>Have fun!</P>
</FONT></BODY>
</HTML>

