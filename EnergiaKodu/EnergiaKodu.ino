#include <SPI.h>
#include <LiquidCrystal.h>
#include <MAX31855.h>

// sicaklik okumak icin pinler tanimlaniyor
const  unsigned  char thermocoupleSO = A7; 
const  unsigned  char thermocoupleCS = A11; 
const  unsigned  char thermocoupleCLK = A10; 

MAX31855  MAX31855(thermocoupleSO, thermocoupleCS, thermocoupleCLK);

//lcd icin pinleri tanimladik
LiquidCrystal lcd(PD_2,PD_3,PB_3,PB_2,PB_6,PB_7);

const int buttonPin = PUSH2; // buton tanimlaniyor
int buttonState = 0; // buton durumu icin deger tanimlaniyor
int stateControl = 0; // false durum ile ilkledik
//ledler tanimlaniyor
#define RED RED_LED
#define GREEN GREEN_LED
#define BLUE BLUE_LED

void setup()
{
  
  // put your setup code here, to run once:
     Serial.begin(9600);
     lcd.begin(16,1);
     
      pinMode(buttonPin, INPUT_PULLUP);
    pinMode(RED, OUTPUT);  
    pinMode(GREEN, OUTPUT);     
    pinMode(BLUE, OUTPUT);
    pinMode(PA_4,OUTPUT); /// calistigini gosteren pinin ilklenmesi
    
}
void loop()
{
  buttonState = digitalRead(buttonPin);
    
    double heating;     
  
 if (buttonState == HIGH){
       if(stateControl == 0){  
          heating = MAX31855.readJunction(CELSIUS);
          Serial.print("Junction temperature: ");
          Serial.print(heating);
          Serial.println(" Degree Celsius");
       }
       else if(stateControl == 1)
       {
          heating = MAX31855.readThermocouple(CELSIUS);
          Serial.print("Thermocouple temperature: ");
          Serial.print(heating);
          Serial.println(" Degree Celsius");
       }
 }
 else if(buttonState == LOW){
        switch(stateControl)
         {
          case 0: stateControl = 1; delay(500); break;
          case 1: stateControl = 0; delay(500); break;
         }
 } 
  if(heating)
  {
          delay(500);
          lcd.setCursor(0,0);
          //LCDye burada yazdiriyor
          lcd.print("temp");
          lcd.print(heating);
          
          if(heating <= 10 )
          {
                 //cyan
                 digitalWrite(RED,LOW);
               digitalWrite(BLUE, HIGH);
               digitalWrite(GREEN,HIGH);
          }
          else if(heating>10 && heating<=20)
           {          //blue
               digitalWrite(RED,LOW);
               digitalWrite(GREEN,LOW);
               digitalWrite(BLUE, HIGH);
             }
           else if( heating>20 && heating<=30)
           {
                 //magenta
               digitalWrite(GREEN,LOW);
               digitalWrite(RED,HIGH);
               digitalWrite(BLUE,HIGH);
           }
           else if(heating>30 && heating<=40)
           {
             //RED
               digitalWrite(GREEN,LOW);
               digitalWrite(RED,HIGH);
               digitalWrite(BLUE,LOW);
           }
           
           //eger okuyamiyorsa extra bir kirmizi isik yanar ve bize calisamadigini anlatir.
         digitalWrite(PA_4,HIGH);
          
          delay(1000);
  }
  else
  {
           //eger sicakligi okuyamazsa extra led yanmaz!
          digitalWrite(PA_4,LOW);
  }
  
  
}
