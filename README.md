# Macros_VBA
## Copiando apenas último e-mail de um loop, tanto em En-US quanto em Pt-BR

## Lógica
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
	

## Como aplicar: 
1.	Enable Macros.
Open Outlook.
From the File menu, click Options --> Trust Center --> Trust Center Settings --> Macro Settings.
Click Ok twice.
 
2.	Adding the Macro to Visual Basic (VB).
Press Alt-F11 to open the VB window.
Right-Click the blank area and choose Insert --> Module, then paste the code in the module Window.


