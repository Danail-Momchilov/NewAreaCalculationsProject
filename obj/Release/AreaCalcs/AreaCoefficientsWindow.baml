   M S B A M L   `   `   `  ���� K  GAreaCalculations, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null3clr-namespace:AreaCalculationsAreaCalculations  -    'AreaCalculations.AreaCoefficientsWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     @d2http://schemas.microsoft.com/expression/blend/2008     0localclr-namespace:AreaCalculations       � SizeToContent$  WidthAndHeight=� � Title$ Area Coefficients�� � 
ResizeMode$ NoResize=� � WindowStyle$ SingleBorderWindow=�
����������#EEE809.���� .+��� .+��� $J��DМоля, въведете всички Area коефициенти:��$��30��$	��300��$
��Left=�$��0,0,0,0q����#EEE809$��Center=���  �� Orientation$ 
Horizontal=�$��
10,0,0,245q�$��30��.+��� $&�� A Coefficient Orientation (Ки)��$��30��$	��250���� -    ��Name$
 ACOR��$��Right=�$��20��$��5,5,0,0q� �� TextWrapping$
 Wrap=�$	��1.0��$	��Top=�$��25���� $ 
Horizontal=�$��10,-460,0,0q�$��30��.+��� $ ��A Coefficient Level (Кв)��$��30��$	��250���� -   $
 ACLE��$��Right=�$��20��$��5,5,0,0q�$
 Wrap=�$	��1.0��$	��Top=�$��25���� $ 
Horizontal=�$��10,-400,0,0q�$��30��.+��� $#��A Coefficient Location (Км)��$��30��$	��250���� -   $
 ACLO��$��Right=�$��20��$��5,5,0,0q�$
 Wrap=�$	��1.0��$	��Top=�$��25���� $ 
Horizontal=�$��10,-340,0,0q�$��30��.+��� $#��A Coefficient Height (Кив)��$��30��$	��250���� -   $
 ACHE��$��Right=�$��20��$��5,5,0,0q�$
 Wrap=�$	��1.0��$	��Top=�$��25���� $ 
Horizontal=�$��10,-280,0,0q�$��30��.+��� $!��A Coefficient Roof (Кпп)��$��30��$	��250���� -   $
 ACRO��$��Right=�$��20��$��5,5,0,0q�$
 Wrap=�$	��1.0��$	��Top=�$��25���� $ 
Horizontal=�$��10,-220,0,0q�$��30��.+��� $$��A Coefficient Special (Кок)��$��30��$	��250���� -   $
 ACSP��$��Right=�$��20��$��5,5,0,0q�$
 Wrap=�$	��1.0��$	��Top=�$��25���� $ 
Horizontal=�$��10,-160,0,0q�$��30��.+��� $ ��A Coefficient Zones (Кк)��$��30��$	��250���� -   $
 ACZO��$��Right=�$��20��$��5,5,0,0q�$
 Wrap=�$	��1.0��$	��Top=�$��25���� $ 
Horizontal=�$��10,-100,0,0q�$��30��.+��� $��A Coefficient Correction��$��30��$	��250���� -   $
 ACCO��$��Right=�$��20��$��5,5,0,0q�$
 Wrap=�$	��1.0��$	��Top=�$��25���� $ 
Horizontal=�$��
10,-40,0,0q�$��30��.+��� $<��6A Coefficient Storage (Ксп) (Склад / Мазе)��$��30��$	��250���� -	   $ ACSTS��$��Right=�$��20��$��5,5,0,0q�$
 Wrap=�$	��0.3��$	��Top=�$��25���� $ 
Horizontal=�$��	10,-5,0,0q�$��30��.+��� $$��A Coefficient Storage (Ксп)��$��30��$	��250���� -
   $
 ACST��$��Right=�$��20��$��5,5,0,0q�$
 Wrap=�$	��1.0��$	��Top=�$��25���� $ 
Horizontal=�$��10,0,0,0q�$��30��.+��� $A��;A Coefficient Garage (Кпг) (Гараж / Паркинг)��$��30��$	��250���� -   $ ACGAP��$��Right=�$��20��$��5,5,0,0q�$
 Wrap=�$	��0.8��$	��Top=�$��25���� $ 
Horizontal=�$��10,0,0,0q�$��30��.+��� $#��A Coefficient Garage (Кпг)��$��30��$	��250���� -   $
 ACGA��$��Right=�$��20��$��5,5,0,0q�$
 Wrap=�$	��1.0��$	��Top=�$��25���� $ 
Horizontal=�$��10,0,0,0q�$��30��.+��� $��Override existing values��$��30��$	��250���� -    ��Name$ overrideCoefficients��$��15��$��15��$��5,0,0,0q� d� 	IsChecked$ FalseY�}�. 	 J� RenderTransformOrigin$	 0.5,0.5&�z�V� .~��� �� �� T� 
 T� X$
 5Q��� -   $��ВЪВЕДИ��$��20��$��0,10,0,0q�$	��280�����#EEE809�� $��20��
������ 