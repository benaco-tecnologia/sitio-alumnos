<%
 dim salida, paso
    
    dim a 'as PagoWeb.PagoWebClase
    'Set a = Server.CreateObject("PagoWeb.InterfacePagoWeb") 'no funciona en asp
    Set a = Server.CreateObject("PagoWeb.PagoWebClase")
    'var a = new ActiveXObject("PagoWeb.PagoWebClase")
    
    
    salida = "sin datos"
    'salida = "b2blCbOiRHN9kCRAMmiUtQENQIwHmW8Ja2JMz2sZIywbbZlJ6b/H8cBcRLxWANLcYVkeVWz8 mvHPymxC2jUEIQoVcWxRI4AffYy4N2orkQxqCrl0p/KXUfbA2Am3KBvHcS2nbenidT5s0hqt D9X//XZONwGG54E7nTepElXGzEg="
    paso = "11111"
    'salida = a.Encripta(paso) 'no funciona
    salida = a.Encriptameloahora(paso)
    
    
    response.Write("*" + salida + "*")
    response.End
    
    
     %>