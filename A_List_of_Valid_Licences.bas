Attribute VB_Name = "A_List_of_Valid_Licences"
Function GetCustomerIDs() As Variant
    ' Returns an array of customer IDs from a dictionary
    Dim dict As Object
    Dim customerArray() As Long
    Dim i As Long

    ' Create a dictionary to map customer IDs to their names
    Set dict = CreateObject("Scripting.Dictionary")
  dict.Add 4209838, "David Shuvalov" 'davidshuvalov@gmail.com
dict.Add 1491687, "Dave Fisher" 'dave@multiwalk.net
dict.Add 2227285, "Shaun Lawman" 'Trading@livelouder.com
dict.Add 4628537, "Christopher Kilgore" 'christopher.a.kilgore@gmail.com
dict.Add 767072, "Al Biddinger" 'albiddinger@gmail.com
dict.Add 1304080, "Zamir & Dunn" 'thequantnetwork@gmail.com
dict.Add 1577327, "Timo Mohnani" 'timo.mohnani@gmail.com; timo.mohnani@protonmail.com
dict.Add 4576263, "Bert Trouwers" 'bert.trouwers@gmail.com
dict.Add 3352708, "Simon Gale" 'simongale@me.com
dict.Add 1438899, "Clay Crandall" 'clayce7@gmail.com
dict.Add 4584703, "Michal Filipkowski" 'michalfilipkowski@yahoo.com
dict.Add 1623185, "Robert Fleming" 'us8lander@gmail.com
dict.Add 982187, "Sean Cooper" 'smcjb@yahoo.com
dict.Add 3534333, "chad ockham" 'chad@ockhamtrading.com
dict.Add 3281546, "Greg Baker" 'gregbaker88@gmail.com
dict.Add 4132757, "Don LaPel" 'lapel.don@gmail.com
dict.Add 2706911, "Philippe Bremard" 'tiger.development.fund@gmail.com
dict.Add 2791698, "Ed Tulauskas" 'tulauskas@gmail.com
dict.Add 874960, "Sanjay Sardana" 'sanjay.sardana@gmail.com
dict.Add 3161289, "Marc Jusseaume" 'marc.jusseaume@icloud.com
dict.Add 1993965, "James Parker" 'dradamas@protonmail.com
dict.Add 1224801, "Gary McOmber " 'mcomber7@gmail.com
dict.Add 1426160, "Daniel Bangert" 'daniel@bangert.com
dict.Add 4305976, "David Aczel" 'ddaczel@icloud.com
dict.Add 4056398, "Niko Heir" 'heir.niko@gmail.com
dict.Add 1447036, "James Welborn" 'jwelbo2004@yahoo.com
dict.Add 828124, "Mark Holland" 'herrmarkholland@gmail.com
dict.Add 4304998, "Nikita Gorbachenko" 'quantadriatic@gmail.com
dict.Add 2749694, "Jayce Nugent" 'jaycenugent@hotmail.com
dict.Add 2697537, "Edwin Shih" 'egshih@gmail.com
dict.Add 4565694, "Victor Stokmans" 'vstokmans@hotmail.com
dict.Add 2546277, "Herman Fuchs" 'herman.fuchs@gmail.com
dict.Add 2210950, "Rajendra Deshpande" 'technicaltrader21@gmail.com
dict.Add 4653662, "Rey Farne" 'rfarne@verizon.net
dict.Add 3923344, "Ujae Kang " 'Ujae Kang <ujae.a.kang@gmail.com>
dict.Add 2986779, "Dave edwards" 'hogface1821@gmail.com
dict.Add 1411262, "Jonas Hellwig" 'Jonas.Hellwig@gmx.net
dict.Add 3144964, "Justin Krick " 'jkrick33@gmail.com
dict.Add 2069314, "Tom Garesche" 'tgaresche@comcast.net
dict.Add 4247492, "Love Englund" 'love@muraya.com
dict.Add 2453776, "Livio Pietroboni" 'livio.pietroboni@hotmail.com
dict.Add 4400422, "Ender Araujo" 'endera2895@gmail.com
dict.Add 2809194, "Dan Omalley" 'danomalley67@gmail.com
dict.Add 3273971, "Covington Creek" 'covingtoncreekvet@gmail.com
dict.Add 1966649, "Ron Mullet" 'rmullet976@aol.com
dict.Add 588649, "James MAZZOLINI" 'jmazz@pacbell.net
dict.Add 4363638, "Seuk Oh" 'osw0309@gmail.com
dict.Add 645309, "John Dorsey" 'jdaz2000@gmail.com
dict.Add 4613075, "Haro Hollertt" 'hhollertt@gmail.com
dict.Add 3213888, "Ryan Williams" 'rlwtrader@yahoo.com
dict.Add 4518976, "Jani Talikka" 'jani.talikka@gmail.com
dict.Add 4363735, "Richard Moore" 'richmoore123@gmail.com
dict.Add 1808839, "Vernon Pratt" 'prattnyu@gmail.com
dict.Add 3870813, "Michal Ko?ousek" 'kodousek.michal@gmail.com
dict.Add 2824178, "Robert Roubey" 'rroubey@gmail.com
dict.Add 3551682, "Venkatesh Yarraguntla" 'venkatesh.yarraguntla@gmail.com
dict.Add 3285215, "Venkatesh Yarraguntla" 'venkatesh.yarraguntla@gmail.com
dict.Add 4744449, "Pete LaDuke" 'pladuke99@gmail.com
dict.Add 3518605, "YOUSSEF OUMANAR" 'oumanary@gmail.com
dict.Add 4230663, "Andreas Savva" 'asavva100@gmail.com
dict.Add 2335131, "Timothy Krull" 'tkinvest@comcast.net
dict.Add 3488626, "Eric Rosko" 'eric_rosko@hotmail.com
dict.Add 4795843, "Denis Smirnov" 'rus101@yahoo.com
dict.Add 3277186, "Thomas Uselton " 'Useltom@gmail.com
dict.Add 4352636, "Anguera Antonio" 'anguera.antonio@gmail.com
dict.Add 4301470, "Miguel Bermejo" 'miguelabs@gmail.com
dict.Add 4760562, "Rohan Patil" 'rohan11188@gmail.com
dict.Add 2139616, "OUrocketman" 'tyler.s.rainey@gmail.com
dict.Add 4877957, "Arin" 'arinm1@icloud.com
dict.Add 4408498, "Arturo Patino" 'arturo.roberto.pat@gmail.com
dict.Add 4874084, "Younes Zerhari" 'zerhari@gmail.com





    ' Resize the array to the number of items in the dictionary
    ReDim customerArray(0 To dict.count - 1)

    ' Fill the array with keys (customer numbers)
    i = 0
    Dim key As Variant
    For Each key In dict.keys
        customerArray(i) = key
        i = i + 1
    Next key

    ' Return the array of customer IDs
    GetCustomerIDs = customerArray
End Function
