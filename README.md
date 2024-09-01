# Birthmails-Auto
Éste pequeño proyecto familiar automatiza el envío de correos electrónicos relacionados con la fecha de cumpleaños de familiares y cercanos; obiamente se puede adaptar a cualquier otro ámbito.

Primero se autentica con credenciales ante Google para pedir permisos y poder loguearse.

Luego lee un googlesheet en donde se va registrando la información de las personas de interés en un DataFrame

El script itera el df para buscar coincidencias en las fechas, si encuentra un cumpleaños para el día de mañana(avisa con un día de anticipación), busca los datos necesarios para escribirlos y enviarlos en el email a los
destinatarios previamente elegidos.


Con solo correr el proceso ya puede realizar todas éstas acciones, por ende, se puede automatizar con el
Programador de Tareas de Windows y colocarle una hora diaria para que ejecute el script.

---

Little project that automates the sending of emails related to birthday dates.