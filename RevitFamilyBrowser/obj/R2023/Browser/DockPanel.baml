   M S B A M L   `   `   `  ���� Q  MzRevitFamilyBrowser, Version=2017.0.0.0, Culture=neutral, PublicKeyToken=nullQ-clr-namespace:zRevitFamilyBrowser.WPF_ClasseszRevitFamilyBrowser.WPF_Classes  /    )zRevitFamilyBrowser.WPF_Classes.DockPanel   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     ?local-clr-namespace:zRevitFamilyBrowser.WPF_Classes     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     -     >�Name$  myDockPanel��
������� � 	Resources    FamilyHeaderStyle �� 
TargetType �� Property �� Value �� 
TargetType 9�Name 9� 
IsExpanded �� Orientation��%�  &	        �� " ��.���  � �� �� " ��.R�9� $ Expander��
�������
�������
���   � .  ���� 
�������$��1q�._��� $	��500��
�������$ 
Horizontal=�.+��� $��5��
��������� $��0��$��0��$��50��$��50��$��80������ *Items[0].img+�� $��5��
��������� $��0��$��0��$	��500��$��5,0,0,0q�$
��Left=�$��Center=����� *Name+.���� .��� .+��� -   	 ��Name$	 
tabControl��$��0,0,0,0q�.@��� $��New Elements��.��� 
�������.+�� 
 � ColumnDefinitions
 �� $��60 ��� $��* ��� $��60 �.+��� -    ��Name$ label��$��0��$��60��$��25��$
��Left=�$	��Top=�$��
Category: ���� $��1��$��25��$��	25,0,20,0q�$	��Top=�.+��� -    ��Name$ label_CategoryName��$��20��$��Stretch=�$��Center=�����  �� ElementName$ myDockPanel�� �� Path$ familyNameFolder� �� UpdateSourceTrigger$ PropertyChanged=����� .R��� $��25������  ��%1   '�       �� " �.���  � ��$ Bottom=��� -   $��2��$��60��$��25��$��Right=�$	��Top=�$��Delete��&   System.Windows.Controls.DataGrid  -     Name$ dataGrid��$��0,25,0,0q�   AutoGenerateColumns .  $��0q�!   EnableColumnVirtualization .     EnableRowVirtualization .     HeadersVisibility$ Column=�   HorizontalGridLinesBrush
 �����   
IsReadOnly .    VerticalGridLinesBrush
 ����� �� 
GroupStyle (  "System.Windows.Controls.GroupStyle     ContainerStyle# [     Panel �� .R�>  8System.Windows.Controls.Primitives.DataGridRowsPresenter     Columns 4  .System.Windows.Controls.DataGridTemplateColumn  ,  &System.Windows.Controls.DataGridColumn   Width5  /System.Windows.Controls.DataGridLengthConverter$ 85    Header$ Image��   CellTemplate �� 0  *System.Windows.Controls.DataGridTextColumn  $ * $
 Name��1  +System.Windows.Controls.DataGridBoundColumn   Binding �� *Name+    ElementStyle  �� .���  � ��$ Center=���  � ��$ 10, 0, 10, 0q��� $��Elements in project��.��� 
�������.+�  -   $ dataGridHistory��$��0,25,0,0q� .  $��0q� .   .  $ Column=�
 ����� . 
 �����   # [   �� .R�     $ 85 $ Image�� ��   $ * $
 Name�� �� *Name+  �� .���  � ��$ Center=���  � ��$ 10, 0, 10, 0q�