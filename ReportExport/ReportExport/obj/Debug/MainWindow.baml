   M S B A M L   `   `   `  ����G  CReportExport, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null+clr-namespace:ReportExportReportExport  5   	       ReportExport.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   ,localclr-namespace:ReportExport     5   	     � Title$  
MainWindow��5   	   $	��300��6   $	��400��6)    � 
ResizeMode$ CanMinimize=�65   .��5	      �    	MainColor#��[  6    � RowDefinitions 5
   
   �� 5      $��50 �6   �� 5      $��* �6   .+�5   
   � $��0��6   ��5      �� 5       �� ImageSource$ /Resources/Header.png��6    j� TileMode$ FlipX=�6A    j� Stretch$
 Fill=�6R    j� 
AlignmentY$	 Top=�6a    j� Viewport$ 950,50,950,50�6r    j� ViewportUnits$ Absolute=�6�   .+�5      �� 	 �� Orientation$	 
Horizontal=�6   .+�5      �� $%��/Resources/betar_logo_white.png��6   $��5,5,15,5q�6A   $
��Left=�6S   �� 5      $��МЕТТЭМ-БЕТАР��6   $��16(�60   
�������$��Center=�6Q   $��Italic$�6l   5      5      �� 5   
   $��1��6   $	 Vertical=�6"   $��Center=�69   $��Center=�6V   .+�5      �� -   5        $��ExelExportBtn��5      $��Экспорт в Exel��6*   $��30��6C   $	��200��6O   $��0,0,0,15q�6[     btnExel#��[ 6m   6�   �� 5      -   5        $��WordExportBtn��5      $��Экспорт в Word��6*   $��30��6C   $	��200��6O     btnWord#��[ 6[   6|   5      5      5      