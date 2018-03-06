# Macros_VBA

String string1, string2, string3
Boolean TemDe = false

Copiar todo o texto do e-mail para string1
Fazer o split de string1 com o delimitador "from: " e salvar em string2

Fazer uma busca em string2 por "De: " 
SE encontrar:
	atribuir true para TemDe
	copiar todo o texto anterior ao "De: " (left) para string3
	passar string3 para a área de transferência
SE não encontrar "De:" (else)
	TemDe continua false
	passar string2 para a área de transferência
	
