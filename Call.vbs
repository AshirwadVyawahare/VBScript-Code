a =2
b=5

abc = add (a,b)  'Works
'add(a,b)         'Do not work
add a,b            'works
call add(a,b)      'works == call statement discards return value by function name

msgbox "Success"

function add (a , b)
 a = b
 add = b
end function