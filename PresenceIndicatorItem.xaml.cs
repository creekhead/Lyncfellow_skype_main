/*=====================================================================
  This file is part of the Microsoft Unified Communications Code Samples.

  Copyright (C) 2012 Microsoft Corporation.  All rights reserved.

This source code is intended only as a supplement to Microsoft
Development Tools and/or on-line documentation.  See these other
materials for detailed information regarding Microsoft code samples.

THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
PARTICULAR PURPOSE.
=====================================================================*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Diagnostics;

namespace LyncFellow
{
    /// <summary>
    /// Interaction logic for PresenceIndicatorItem.xaml
    /// </summary>
    public partial class PresenceIndicatorItem : UserControl
    {
        public string SourceUri { set { presenceIndicator1.Source = value; } }
        

        public PresenceIndicatorItem()
        {
            //MessageBox.Show("PresenceIndicatorItem");
            InitializeComponent();
        }
    }
}
