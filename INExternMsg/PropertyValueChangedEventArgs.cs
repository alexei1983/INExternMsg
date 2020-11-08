using System;

namespace org.goodspace.Utils.ImageNow {

  /// <summary>
  /// Represents event data associated with changing the value of an existing property 
  /// in an <see cref="INExternMsg"/>.
  /// </summary>
  public class PropertyValueChangedEventArgs : EventArgs {

    /// <summary>
    /// The name of the property affected by the change.
    /// </summary>
    public string PropertyName { get; set; }

    /// <summary>
    /// The value of the property after the change occurred.
    /// </summary>
    public string NewPropertyValue { get; set; }

    /// <summary>
    /// The value of the property before the change occurred.
    /// </summary>
    public string OldPropertyValue { get; set; }
  }
}
