
namespace org.goodspace.Utils.ImageNow {

  #region Enums

  /// <summary>
  /// INExternMsg message direction.
  /// </summary>
  public enum ExternMsgDirection {

    /// <summary>
    /// Outbound message.
    /// </summary>
    Outbound = 2,

    /// <summary>
    /// Inbound message.
    /// </summary>
    Inbound = 1,

    /// <summary>
    /// Message direction is undefined.
    /// </summary>
    Undefined = 0,
  }

  /// <summary>
  /// INExternMsg message status.
  /// </summary>
  public enum ExternMsgStatus {

    /// <summary>
    /// Message status is undefined.
    /// </summary>
    Undefined = 0,

    /// <summary>
    /// Status is NEW.
    /// </summary>
    New = 1,

    /// <summary>
    /// Status is PROCESSING.
    /// </summary>
    Processing = 2,

    /// <summary>
    /// Status is COMPLETE.
    /// </summary>
    Complete = 3,

    /// <summary>
    /// Status is ERROR.
    /// </summary>
    Error = 4,

    /// <summary>
    /// Status is LOCKED.
    /// </summary>
    Locked = 999,
  }

  /// <summary>
  /// INExternMsg property type.
  /// </summary>
  public enum ExternMsgPropType {

    /// <summary>
    /// Property type is undefined.
    /// </summary>
    Undefined = 0,
  }

  #endregion
}
