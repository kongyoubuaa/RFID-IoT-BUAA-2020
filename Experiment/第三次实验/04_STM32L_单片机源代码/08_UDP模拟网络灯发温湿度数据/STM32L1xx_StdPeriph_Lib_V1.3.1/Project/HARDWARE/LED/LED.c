#include "led.h"
void LED_Init(void)
{		
    GPIO_InitTypeDef   GPIO_InitStructure;
    RCC_AHBPeriphClockCmd(RCC_AHBPeriph_GPIOA, ENABLE);	 	
    GPIO_InitStructure.GPIO_Mode = GPIO_Mode_OUT;
    GPIO_InitStructure.GPIO_Pin = GPIO_Pin_1|GPIO_Pin_15;//����PA1 PA15Ϊ���
    GPIO_Init(GPIOA, &GPIO_InitStructure);
} 