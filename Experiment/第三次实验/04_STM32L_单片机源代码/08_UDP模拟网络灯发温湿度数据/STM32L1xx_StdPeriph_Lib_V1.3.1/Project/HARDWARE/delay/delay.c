#include "delay.h"

static unsigned char  fac_us=0;//usʹ��
static u16 fac_ms=0;//msʹ��
void delay_init()	 
{


	SysTick_CLKSourceConfig(SysTick_CLKSource_HCLK_Div8);	//ѡ���ⲿʱ��  HCLK/8
	fac_us=SystemCoreClock/8000000;	//Ϊϵͳʱ�ӵ�1/8  4
	fac_ms=(u16)fac_us*1000;
}								    

void delay_ms(u16 nms)
{	 		  	  
	 uint32_t ui_tmp = 0x00;  
    SysTick->LOAD = nms * fac_ms;  
    SysTick->VAL = 0x00;  
    SysTick->CTRL = 0x01;  
      
    do  
    {  
        ui_tmp = SysTick->CTRL;  
    }while((ui_tmp & 0x01) && (!(ui_tmp & (1 << 16))));  
      
    SysTick->CTRL = 0x00;  
    SysTick->VAL = 0x00;      
} 




void delay_us(u16 nus)
{	 		  	  
	 uint32_t ui_tmp = 0x00;  
    SysTick->LOAD = nus * fac_us;  
    SysTick->VAL = 0x00;  
    SysTick->CTRL = 0x01;  
      
    do  
    {  
        ui_tmp = SysTick->CTRL;  
    }while((ui_tmp & 0x01) && (!(ui_tmp & (1 << 16))));  
      
    SysTick->CTRL = 0x00;  
    SysTick->VAL = 0x00;      
} 





void Delay(u16 nms)
{	 		  	  
	 uint32_t ui_tmp = 0x00;  
    SysTick->LOAD = nms * fac_ms;  
    SysTick->VAL = 0x00;  
    SysTick->CTRL = 0x01;  
      
    do  
    {  
        ui_tmp = SysTick->CTRL;  
    }while((ui_tmp & 0x01) && (!(ui_tmp & (1 << 16))));  
      
    SysTick->CTRL = 0x00;  
    SysTick->VAL = 0x00;   	    
} 































