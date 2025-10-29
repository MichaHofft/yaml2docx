# Translate

translated via: et-Content .\aas-grammar-iec-63278-5.ebnf | docker run --rm -i -v "${PWD}:/data" kgt -l iso-ebnf -e svg

# Convert BNF to ISO ebnf

```
replace: ^(.{2,999})$
with: $1 ;

replace: ::=
with: =

replace: <(\w+)>
with: $1,

replace: , =
with: = 

replace: , \|
with:  |

replace: , ;
with:  ;

replace: ("[^"]+")
with: $1, 

replace: ,  ;
with:  ; 

1-by-1 replace: \(([^)]*)\)\?
with: {$1}

replace: ,\s+}
with:  }

replace: ,\s*=
replace:  =

replace: ,\s+\)
with:  \),

```