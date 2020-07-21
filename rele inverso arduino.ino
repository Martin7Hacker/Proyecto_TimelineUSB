
// programita para utilizar una aplicación .exe con Arduino
// a lo que el relé es de logica invertida creo el comparador comp de formato entero.
int led=2; // dispositivo a activar
int comp=0;// comparador de inicio para que el relé quede apagado siempre.
byte dato; // dato por el USB del puerto de comunicación de lectura y escritura.

// configuración de inicio.
void setup(){
  pinMode(led,OUTPUT);// tipo de modo de pin Salida
  Serial.begin(9600); // inicializacion de baudios frecuencia del puerto 9600.
}

// repeticion infinita
void loop(){
  if (comp==0){
   digitalWrite(led,HIGH); // se mantiene en estado apagado el relé porque es de logica invertida
  }

// lectura del puerto //
  dato=Serial.read();// léer el dato que me da la aplicacion .exe
  if(dato=='0'){
    digitalWrite(led,HIGH);// el Led se mantiene encendido
    comp=0;
  }

  if(dato=='1'){
    digitalWrite(led,LOW);// el led se mantiene apagado
    comp=1;
  }
  
  delay(77); // tiempo de descanso del relé.
  
}

