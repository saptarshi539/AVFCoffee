﻿using CoffeeCore.DTO;
using System;
using System.Collections.Generic;
using System.Text;

namespace CoffeeCore.Interfaces
{
    public interface ITechnicianFlexcelSum
    {
        TechnicianLoginInfoDTO GetUserMetrics();
        void saveUserMetrics(String[] data);
        void saveUserAdvancedInputs(ChartInputAdvancedDTO advancedInputs);
        Dictionary<string, object> getInputs();
        ChartInputAdvancedDTO getInputValues();
    }
}
