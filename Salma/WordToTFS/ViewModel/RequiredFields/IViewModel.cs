namespace WordToTFS.ViewModel
{
    /// <summary>
    /// Base class for required fields
    /// </summary>
    public interface IViewModel
    {
        string GetName();
        object GetValue();
        bool IsNumeric { get; }
    }
}