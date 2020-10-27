#include <stm32l1xx.h>
 void LED_Init(void);
#define LEDMCU_1   GPIO_SetBits(GPIOA,GPIO_Pin_1)
#define LEDMCU_0   GPIO_ResetBits(GPIOA,GPIO_Pin_1)
#define LEDNET_1   GPIO_SetBits(GPIOA,GPIO_Pin_15)
#define LEDNET_0   GPIO_ResetBits(GPIOA,GPIO_Pin_15)