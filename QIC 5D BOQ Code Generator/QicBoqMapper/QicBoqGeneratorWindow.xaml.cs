using System;
using System.Windows;

namespace QicBoqMapper
{
    public partial class QicBoqGeneratorWindow : Window
    {
        private readonly QicBoqGeneratorViewModel _viewModel;

        public QicBoqGeneratorWindow(QicBoqGeneratorViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = _viewModel;
        }

        private void SelectAllCategories_Click(object sender, RoutedEventArgs e)
        {
            if (_viewModel?.CategoryItems == null) return;
            foreach (var item in _viewModel.CategoryItems)
            {
                item.IsSelected = true;
            }
        }

        private void SelectNoCategories_Click(object sender, RoutedEventArgs e)
        {
            if (_viewModel?.CategoryItems == null) return;
            foreach (var item in _viewModel.CategoryItems)
            {
                item.IsSelected = false;
            }
        }
    }
}
