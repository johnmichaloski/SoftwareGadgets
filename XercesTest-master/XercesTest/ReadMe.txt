   /*
    Create an empty DOMDocument. When an empty document is created specify the 
         namespace of the document (optional)
    and also the root node. Both of these values can be any valid text 
         (ie. no spaces or weird characters).
    The "root node" in this example is "Hello_World". (XML is case-sensitive) 
         This root node *owns everything* beneath it, just as if it were a separate
         object... in fact, it is. It's defined
    by the 'DOMNode' class.
    */
    p_DOMDocument = p_DOMImplementation->createDocument(0, L"Hello_World", 0);

    /*
    The output of the code above will produce the following:

        <?xml version="1.0" encoding="UTF-8" standalone="no" ?>
        <Hello_World/>

    */