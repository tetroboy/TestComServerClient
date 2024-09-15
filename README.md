# TestComServerClient
So u need to build testCom project(sln) After that u need to check in registry if clsid and inrface is correctly registered
than u have to build Typelib.cpp file and look if typelib registered well
and than if all ok u can build client code 

Your ClsID have to have inside a key named typelib with default value of GUID of typelib
and also key named version with Default value 1.0
look something like that
(Guid of clsid)
-typelib
-version
-localserver32
