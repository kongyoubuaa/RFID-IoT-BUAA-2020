#ifndef __DHT11_H
#define __DHT11_H 
#include <stm32l1xx.h>
#include "delay.h"
////IO��������											   
//#define	DHT11_DQ_OUT GPIO_SetBits(GPIOB,GPIO_Pin_7) //���ݶ˿�	PA7
#define	DHT11_DQ_IN  GPIO_ReadInputDataBit(GPIOA,GPIO_Pin_7)  //���ݶ˿�	PA7

u8 DHT11_Init(void);//��ʼ��DHT11
u8 DHT11_Read_Data(u8 *temp,u8 *humi);//��ȡ��ʪ��
u8 DHT11_Read_Byte(void);//����һ���ֽ�
u8 DHT11_Read_Bit(void);//����һ��λ
u8 DHT11_Check(void);//����Ƿ����DHT11
void DHT11_Rst(void);//��λDHT11    
void DHT11_IO_IN(void);
void delay_us(uint32_t value);
void DHT11_Rst(void)   ;
#endif















