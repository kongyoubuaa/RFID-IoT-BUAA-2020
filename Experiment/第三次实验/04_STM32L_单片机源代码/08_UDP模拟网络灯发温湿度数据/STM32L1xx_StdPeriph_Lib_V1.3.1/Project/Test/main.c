/**
  ******************************************************************************
  * @file    Project/STM32L1xx_StdPeriph_Templates/main.c 
  * @author  MCD Application Team
  * @version V1.2.0
  * @date    16-May-2014
  * @brief   Main program body
  ******************************************************************************
  * @attention
  *
  * <h2><center>&copy; COPYRIGHT 2014 STMicroelectronics</center></h2>
  *
  * Licensed under MCD-ST Liberty SW License Agreement V2, (the "License");
  * You may not use this file except in compliance with the License.
  * You may obtain a copy of the License at:
  *
  *        http://www.st.com/software_license_agreement_liberty_v2
  *
  * Unless required by applicable law or agreed to in writing, software 
  * distributed under the License is distributed on an "AS IS" BASIS, 
  * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
  * See the License for the specific language governing permissions and
  * limitations under the License.
  *
  ******************************************************************************
  */

/* Includes ------------------------------------------------------------------*/
#include "main.h"
#include "usart.h"
#include "timer.h"
#include "bc95.h"
#include "led.h"
#include "delay.h"
#include "dht11.h"
#include "string.h"
#include "stdlib.h"
GPIO_InitTypeDef GPIO_InitStructure;
static __IO uint32_t TimingDelay;
u8 temp,humi;
/* Private function prototypes -----------------------------------------------*/
/** @addtogroup Template_Project
  * @{
  */

/************通过转发网关的形式将数据发到个人服务器端，电信受IP限制，移动NB卡不限制**************/
/**
  * @brief  Main program.
  * @param  None
  * @retval None
  */
	
int main(void)
{
    u8 sendata[]="3CCFED54901FFF1835";
	  u8 i;
    delay_init();	
    LED_Init();
    while(DHT11_Init());//初始化DHT11
	
    uart_init(9600);
    uart3_init(9600);
    TIM4_Int_Init(4999,3199);//500ms一次中断
	  Uart1_SendStr("BC95-INIT\n");
    BC95_Init();
	  Uart1_SendStr("BC95-PDPACT\n");	
    BC95_PDPACT();
	  Uart1_SendStr("BC95-ConUDP\n");		
    BC95_ConUDP();
	
  while (1)
  {
		DHT11_Read_Data(&temp,&humi);//读取温湿度数据
		sendata[14]=temp/10+0x30;
		sendata[15]=temp%10+0x30;
		sendata[16]=humi/10+0x30;
		sendata[17]=humi%10+0x30;//转成字符形式
	  Uart1_SendStr(sendata);	
		BC95_Senddata("9",sendata);//转发方式网关，我们提供的电信卡。如果用移动NB卡不限制IP。
        delay_ms(500);
        BC95_RECData();//接收下发的数据 
      
  }
}




#ifdef  USE_FULL_ASSERT

/**
  * @brief  Reports the name of the source file and the source line number
  *         where the assert_param error has occurred.
  * @param  file: pointer to the source file name
  * @param  line: assert_param error line source number
  * @retval None
  */
void assert_failed(uint8_t* file, uint32_t line)
{ 
  /* User can add his own implementation to report the file name and line number,
     ex: printf("Wrong parameters value: file %s on line %d\r\n", file, line) */

  /* Infinite loop */
  while (1)
  {
  }
}
#endif

/**
  * @}
  */


/************************ (C) COPYRIGHT STMicroelectronics *****END OF FILE****/
