using System;
using Xunit;
using NSubstitute;
using src;
using src.Presenters;
using src.Services;
using src.Factories;
using src.Models;
using src.Events;
using System.Collections.Generic;
using System.Linq;

namespace Test
{
    public class MainFormPresenterTests
    {
        private readonly MainFormPresenter<IMainFormView> _sut;
        private readonly IMainFormView _view = Substitute.For<IMainFormView>();
        private readonly IOfficeAppProvider _provider = Substitute.For<IOfficeAppProvider>();

        public MainFormPresenterTests()
        {
            _sut = new MainFormPresenter<IMainFormView>(_provider);
            _sut.View = _view;
        }

        [Fact]
        public void HandleAppClosed_RemovesApplicationFromPresenterAndView()
        {
            // Arrange
            _sut.OpenApps.Add(new OfficeApp(OfficeAppType.Excel, "testExcel"));
            _view.ListOfOpenOfficeApplications = new System.Windows.Forms.ListBox();
            _view.ListOfOpenOfficeApplications.Items.Add("Excel - testExcel");
            var eventArgs = new AppClosedEventArgs("testExcel", OfficeAppType.Excel);

            // Act
            _sut.HandleAppClosed(this, eventArgs);

            // Assert
            Assert.Empty(_sut.OpenApps);
            Assert.Empty(_view.ListOfOpenOfficeApplications.Items);
        }

        [Fact]
        public void HandleAppOpened_AddsNewApplicationToPresenterAndView()
        {
            // Arrange
            _provider.FetchNewWordApplication(0).Returns(new OfficeApp(OfficeAppType.Word, "testWord"));
            _view.ListOfOpenOfficeApplications = new System.Windows.Forms.ListBox();
            _sut.OpenApps.Add(new OfficeApp(OfficeAppType.PowerPoint, "testPp"));

            // Act
            _sut.HandleAppOpened(this, new AppOpenedEventArgs(OfficeAppType.Word));

            // Assert
            Assert.True(_sut.OpenApps.Count == 2);
            Assert.Equal("testWord", _sut.OpenApps[1].Name);
            Assert.True(_view.ListOfOpenOfficeApplications.Items.Count == 1);
        }

        [Fact]
        public void AddAppsOnStartup_PopulatesOpenAppsList()
        {
            // Arrange
            _provider.TryFetchPowerPointApplicationsOnStartup().Returns(Enumerable.Repeat(new OfficeApp(OfficeAppType.PowerPoint, "testPP"), 1));
            _provider.TryFetchExcelApplicationsOnStartup().Returns(Enumerable.Repeat(new OfficeApp(OfficeAppType.Excel, "testXl"), 1));
            _provider.TryFetchWordApplicationsOnStartup().Returns(Enumerable.Repeat(new OfficeApp(OfficeAppType.Word, "testWord"), 1));
            _view.ListOfOpenOfficeApplications = new System.Windows.Forms.ListBox();

            // Act
            _sut.AddAppsOnStartup(_provider.TryFetchPowerPointApplicationsOnStartup);
            _sut.AddAppsOnStartup(_provider.TryFetchExcelApplicationsOnStartup);
            _sut.AddAppsOnStartup(_provider.TryFetchWordApplicationsOnStartup);

            // Assert
            Assert.True(_sut.OpenApps.Count == 3);
            Assert.True(_view.ListOfOpenOfficeApplications.Items.Count == 3);
        }

        [Fact]
        public void ChangeAutoSaveFrequency_UpdatesAutoSaveFrequencyCorrectly()
        {
            // Arrange
            _view.DropDownAutoSaveFrequencies = new System.Windows.Forms.ListBox();
            _view.DropDownAutoSaveFrequencies.Items.AddRange(new object[] {
            "10 Minutes",
            "30 Minute",
            "1 Hour"});
            _view.DropDownAutoSaveFrequencies.SelectedIndex = 0;

            // Act
            _sut.ChangeAutoSaveFrequency();

            // Assert
            Assert.Equal(10000, _sut.CurrentAutoSaveFrequency);
        }

        [Fact]
        public void SubscribeToViewEvents_EventHandlersTriggeredWhenViewEventsFired()
        {
            // Arrange
            bool saveButtonClickEventFired = false;
            bool autoSaveFrequencySelectedEventFired = false;
            bool saveSelectedButtonClickEventFired = false;

            _view.SaveAllButton = new System.Windows.Forms.Button();
            _view.SaveSelectedButton = new System.Windows.Forms.Button();
            _view.DropDownAutoSaveFrequencies = new System.Windows.Forms.ListBox();
            _view.DropDownAutoSaveFrequencies.Items.AddRange(new object[] {
            "10 Minutes",
            "30 Minute",
            "1 Hour"});

            _view.SaveAllButton.Click += (s, e) => saveButtonClickEventFired = true;
            _view.DropDownAutoSaveFrequencies.SelectedIndexChanged += (s, e) => autoSaveFrequencySelectedEventFired = true;
            _view.SaveSelectedButton.Click += (s, e) => saveSelectedButtonClickEventFired = true;

            // Act
            _view.SaveAllButton.PerformClick();
            _view.SaveSelectedButton.PerformClick();
            _view.DropDownAutoSaveFrequencies.SelectedIndex = 1;

            // Assert
            Assert.True(saveButtonClickEventFired);
            Assert.True(autoSaveFrequencySelectedEventFired);
            Assert.True(saveSelectedButtonClickEventFired);
        }
    }
}