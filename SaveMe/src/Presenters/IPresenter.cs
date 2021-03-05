using System;
using System.Collections.Generic;
using System.Text;

namespace src.Presenters
{
    public interface IPresenter<TView>
    {
        TView View { get; set; }

        void SubscribeToViewEvents();
    }
}
