using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace org.goodspace.Utils.ImageNow {

  /// <summary>
  /// Represents an INExternMsg object in the ImageNow application and database.
  /// </summary>
  public class INExternMsg : IFormattable, IEquatable<INExternMsg> {

    #region Operators

    /// <summary>
    /// Returns a value indicating whether the specified <see cref="INExternMsg"/> 
    /// objects represent the same value.
    /// </summary>
    /// <param name="left"><see cref="INExternMsg"/> object on the left side.</param>
    /// <param name="right"><see cref="INExternMsg"/> object on the right side.</param>
    /// <returns><c>True</c> if the <see cref="INExternMsg"/> objects are equal, else <c>false</c>.</returns>
    public static bool operator ==(INExternMsg left, INExternMsg right)
    {

      if (left is null)
        return right is null;

      return left.Equals(right);
    }

    /// <summary>
    /// Returns a value indicating whether the specified <see cref="INExternMsg"/> 
    /// objects do not represent the same value.
    /// </summary>
    /// <param name="left"><see cref="INExternMsg"/> object on the left side.</param>
    /// <param name="right"><see cref="INExternMsg"/> object on the right side.</param>
    /// <returns><c>True</c> if the <see cref="INExternMsg"/> objects are not equal, else <c>false</c>.</returns>
    public static bool operator !=(INExternMsg left, INExternMsg right)
    {

      return !(left == right);
    }

    #endregion

    #region Public events

    /// <summary>
    /// Event raised when the current <see cref="INExternMsg"/> object is sent.
    /// </summary>
    public event EventHandler MessageSent;

    /// <summary>
    /// Event raised when a property value changes in the current <see cref="INExternMsg"/> object.
    /// </summary>
    public event EventHandler<PropertyValueChangedEventArgs> PropertyValueChanged;

    /// <summary>
    /// Event raised when a property is added to the current <see cref="INExternMsg"/> object.
    /// </summary>
    public event EventHandler<PropertyAddedEventArgs> PropertyAdded;

    #endregion

    #region Public properties

    /// <summary>
    /// The message ID.
    /// </summary>
    public string MessageId
    {
      get
      {
        return msgId ?? string.Empty;
      }

      set
      {
        if (!string.IsNullOrEmpty(value))
        {
          if (value.Length > INExternMsgHelper.MAX_MESSAGE_ID_LEN)
            throw new ArgumentException(
                string.Format(
                    "Message ID cannot be more than {0} characters in length.",
                        INExternMsgHelper.MAX_MESSAGE_ID_LEN));
        }

        msgId = value;
      }
    }

    /// <summary>
    /// Start time of the message.
    /// </summary>
    public DateTime StartTime
    {
      get
      {
        return startTime ?? INExternMsgHelper.IMAGENOW_DEFAULT_DATE;
      }

      set
      {
        if (value < INExternMsgHelper.IMAGENOW_DEFAULT_DATE)
          throw new ArgumentException(
              string.Format("Start time cannot be less than ImageNow default date {0}.",
              INExternMsgHelper.IMAGENOW_DEFAULT_DATE.ToString("G", CultureInfo.CurrentCulture)));

        startTime = value;
      }
    }

    /// <summary>
    /// End time of the message.
    /// </summary>
    public DateTime? EndTime
    {
      get
      {
        return endTime;
      }

      private set
      {
        if (value != null)
        {
          if (value < INExternMsgHelper.IMAGENOW_DEFAULT_DATE)
            throw new ArgumentException(
                string.Format("End time cannot be less than ImageNow default date {0}.",
                INExternMsgHelper.IMAGENOW_DEFAULT_DATE.ToString("G", CultureInfo.CurrentCulture)));
        }

        endTime = value;
      }
    }

    /// <summary>
    /// The message direction.
    /// </summary>
    public ExternMsgDirection Direction { get; set; }

    /// <summary>
    /// The message status.
    /// </summary>
    public ExternMsgStatus Status { get; set; }

    /// <summary>
    /// The message type.
    /// </summary>
    public string Type
    {
      get
      {
        return msgType ?? string.Empty;
      }

      set
      {
        if (!string.IsNullOrEmpty(value))
        {
          if (value.Length > INExternMsgHelper.MAX_MSG_TYPE_LEN)
            throw new ArgumentException(
                string.Format(
                    "Message type cannot be more than {0} characters in length.",
                        INExternMsgHelper.MAX_MSG_TYPE_LEN));
        }

        msgType = value;
      }
    }

    /// <summary>
    /// The message name.
    /// </summary>
    public string Name
    {
      get
      {
        return msgName ?? string.Empty;
      }

      set
      {
        if (!string.IsNullOrEmpty(value))
        {
          if (value.Length > INExternMsgHelper.MAX_MSG_NAME_LEN)
            throw new ArgumentException(
                string.Format(
                    "Message name cannot be more than {0} characters in length.",
                        INExternMsgHelper.MAX_MSG_NAME_LEN));
        }

        msgName = value;
      }
    }

    /// <summary>
    /// The number of properties defined in the <see cref="INExternMsg"/>.
    /// </summary>
    public int PropertyCount
    {
      get
      {
        return properties?.Count ?? 0;
      }
    }

    /// <summary>
    /// Gets or sets the value of the specified message property.
    /// </summary>
    /// <param name="propertyName">Name of the property to get or set.</param>
    /// <returns>Value of the property as a <see cref="string"/>.</returns>
    public string this[string propertyName]
    {
      get
      {
        return GetProperty(propertyName);
      }

      set
      {
        SetProperty(propertyName, value, ExternMsgPropType.Undefined);
      }
    }

    /// <summary>
    /// Determines how <see cref="string"/> values are compared in the 
    /// current <see cref="INExternMsg"/>.
    /// </summary>
    public StringComparison StringComparison { get; set; } = StringComparison.InvariantCultureIgnoreCase;

    #endregion

    #region Constructors

    /// <summary>
    /// Creates a new instance of the <see cref="INExternMsg"/> class.
    /// </summary>
    public INExternMsg()
    {
      properties = new Dictionary<string, string>();
    }

    /// <summary>
    /// Creates a new instance of the <see cref="INExternMsg"/> class using the 
    /// specified message type, message name, and connection string.
    /// </summary>
    /// <param name="type">The message type.</param>
    /// <param name="name">The message name.</param>
    public INExternMsg(string type,
                       string name) : this()
    {
      Name = name;
      Type = type;
      Status = ExternMsgStatus.New;
      Direction = ExternMsgDirection.Inbound;
      StartTime = DateTime.Now;
    }

    /// <summary>
    /// Creates a new instance of the <see cref="INExternMsg"/> class using the specified 
    /// message type, message name, connection string, and message status. If the 
    /// message status is New, the message is prepared for sending by setting the 
    /// start time to the current date and time and the direction to Inbound.
    /// </summary>
    /// <param name="type">The message type</param>
    /// <param name="name">The message name</param>
    /// <param name="connectionString">The ImageNow database connection string</param>
    /// <param name="status">The message status</param>
    public INExternMsg(string type,
                       string name,
                       ExternMsgStatus status) : this()
    {
      Type = type;
      Name = name;
      Status = status;

      if (status == ExternMsgStatus.New)
      {
        StartTime = DateTime.Now;
        Direction = ExternMsgDirection.Inbound;
      }
    }

    /// <summary>
    /// Creates a new instance of the <see cref="INExternMsg"/> class using the 
    /// specified values.
    /// </summary>
    /// <param name="type">The message type.</param>
    /// <param name="name">The message name.</param>
    /// <param name="direction">The message direction.</param>
    /// <param name="status">The message status.</param>
    /// <param name="startTime">Start time of the message.</param>
    /// <param name="endTime">End time of the message.</param>
    /// <param name="stringComparison"></param>
    public INExternMsg(string type,
                       string name,
                       ExternMsgDirection direction,
                       ExternMsgStatus status,
                       DateTime startTime,
                       DateTime? endTime = null,
                       StringComparison stringComparison = StringComparison.InvariantCultureIgnoreCase) : this()
    {
      Type = type;
      Name = name;
      Direction = direction;
      Status = status;
      StartTime = startTime;
      EndTime = endTime;
      StringComparison = stringComparison;
    }

    /// <summary>
    /// Creates a new instance of the <see cref="INExternMsg"/> class using the specified 
    /// message type and message name. The message properties in the array are added to
    /// the INExternMsg with null values, and the message is prepared for sending by 
    /// setting the start time to the current date and time, the direction to <see cref="ExternMsgDirection.Inbound"/>, 
    /// and the status to <see cref="ExternMsgStatus.New"/>.
    /// </summary>
    /// <param name="type">The message type.</param>
    /// <param name="name">The message name.</param>
    /// <param name="propertyNames">Array of property names to add to the message.</param>
    public INExternMsg(string type,
                       string name,
                       params string[] propertyNames) : this()
    {
      Type = type;
      Name = name;
      StartTime = DateTime.Now;
      Direction = ExternMsgDirection.Inbound;
      Status = ExternMsgStatus.New;

      if (propertyNames != null)
        foreach (var p in propertyNames)
          if (!string.IsNullOrWhiteSpace(p))
            AddProperty(p, ExternMsgPropType.Undefined);
    }

    /// <summary>
    /// Creates a new instance of the <see cref="INExternMsg"/> class using the specified
    /// <see cref="INExternMsg"/> object as a template. The message type, message name,
    /// connection string and property names are copied into the new 
    /// <see cref="INExternMsg"/> object, and the message is prepared for sending by 
    /// setting the start time to the current date and time, the direction to <see cref="ExternMsgDirection.Inbound"/>, 
    /// and the status to <see cref="ExternMsgStatus.New"/>.
    /// </summary>
    /// <param name="messageTemplate"><see cref="INExternMsg"/> object to use as a 
    /// template for a new <see cref="INExternMsg"/> object.</param>
    public INExternMsg(INExternMsg messageTemplate) : this()
    {
      if (messageTemplate != null)
      {

        Type = messageTemplate.Type;
        Name = messageTemplate.Name;
        StartTime = DateTime.Now;
        Direction = ExternMsgDirection.Inbound;
        Status = ExternMsgStatus.New;
        StringComparison = messageTemplate.StringComparison;

        // Copy the property names
        foreach (var p in messageTemplate.GetPropertyNames())
          AddProperty(p, ExternMsgPropType.Undefined);
      }
    }

    #endregion

    #region Public instance methods

    /// <summary>
    /// Adds a property to the current <see cref="INExternMsg"/> with the specified name 
    /// and a null value.
    /// </summary>
    /// <param name="propertyName">Name of the property to add.</param>
    /// <param name="propertyType">Type of the property to add.</param>
    public void AddProperty(string propertyName,
                            ExternMsgPropType propertyType = ExternMsgPropType.Undefined)
    {
      AddProperty(propertyName, null, propertyType);
    }

    /// <summary>
    /// Adds a property to the current <see cref="INExternMsg"/> with the specified 
    /// name and value.
    /// </summary>
    /// <param name="propertyName">Name of the property to add.</param>
    /// <param name="propertyValue">Value of the property to add.</param>
    /// <param name="propertyType">Type of the property to add.</param>
    public void AddProperty(string propertyName,
                            string propertyValue,
                            ExternMsgPropType propertyType = ExternMsgPropType.Undefined)
    {
      ValidatePropNameAndValue(propertyName, propertyValue, true);

      if (!InternalContains(propertyName, out string definedPropertyName))
      {
        properties.Add(propertyName, propertyValue);

        OnPropertyAdded(propertyName, propertyValue);
      }
      else
      {
        string propVal = GetProperty(definedPropertyName);

        throw new ArgumentException(
            string.Format(
            "Property name '{0}' already exists with value '{1}'",
            propertyName, propVal ?? "null"));
      }
    }

    /// <summary>
    /// Sets the value of the specified property in the current <see cref="INExternMsg"/>.
    /// </summary>
    /// <param name="propertyName">Name of the property to set.</param>
    /// <param name="propertyValue">Value of the property to set.</param>
    /// <param name="propertyType">Type of the property to set.</param>
    public void SetProperty(string propertyName,
                            string propertyValue,
                            ExternMsgPropType propertyType = ExternMsgPropType.Undefined)
    {
      ValidatePropNameAndValue(propertyName, propertyValue, true);

      if (InternalContains(propertyName, out string definedPropertyName))
      {
        var currentValue = GetProperty(definedPropertyName);

        var valueChanged = false;

        if (currentValue == null && propertyValue != null ||
            propertyValue == null && currentValue != null ||
            !(currentValue ?? string.Empty).Equals(propertyValue ?? string.Empty, StringComparison))
          valueChanged = true;

        properties[definedPropertyName] = propertyValue;

        if (valueChanged)
          OnPropertyValueChanged(definedPropertyName, currentValue, propertyValue);
      }
      else
        AddProperty(propertyName, propertyValue, propertyType);
    }

    /// <summary>
    /// Retrieves the value of the specified property in the current <see cref="INExternMsg"/>.
    /// </summary>
    /// <param name="propertyName">Name of the property to retrieve.</param>
    /// <returns>Value of the property as a <see cref="string"/>.</returns>
    public string GetProperty(string propertyName)
    {
      ValidatePropNameAndValue(propertyName, string.Empty, false);

      if (!InternalContains(propertyName, out string definedPropertyName))
        throw new ArgumentException(
            string.Format("Property '{0}' is not defined.", propertyName));

      return properties[definedPropertyName];
    }

    /// <summary>
    /// Retrieves an enumerable list of the property names defined in the 
    /// current <see cref="INExternMsg"/>.
    /// </summary>
    /// <returns>An enumerable list of the property names defined in the 
    /// current <see cref="INExternMsg"/>.</returns>
    public IEnumerable<string> GetPropertyNames()
    {
      if (properties != null)
      {
        foreach (var p in properties.Keys)
          yield return p;
      }
    }

    /// <summary>
    /// Determines whether or not the current <see cref="INExternMsg"/> contains the 
    /// specified property name.
    /// </summary>
    /// <param name="propertyName">Name of the property to check for existence.</param>
    /// <returns><c>True</c> if the current <see cref="INExternMsg"/> contains the specified 
    /// property name, else <c>false</c>.</returns>
    public bool HasProperty(string propertyName)
    {
      ValidatePropNameAndValue(propertyName, null, false);
      return InternalContains(propertyName, out string _);
    }

    /// <summary>
    /// Clears the properties and associated values in the current <see cref="INExternMsg"/>.
    /// </summary>
    public void ClearProperties()
    {
      if (properties == null)
        return;

      properties.Clear();
    }

    /// <summary>
    /// Clears the values of the properties in the current <see cref="INExternMsg"/>.
    /// </summary>
    public void ClearPropertyValues()
    {
      if (properties == null)
        return;

      foreach (var prop in GetPropertyNames().ToList())
        properties[prop] = null;
    }

    /// <summary>
    /// Updates the status of the current outbound <see cref="INExternMsg"/> in the database.
    /// </summary>
    /// <param name="status">The message status to set in the database.</param>
    /// <returns><c>True</c> if the message status was successfully updated 
    /// in the database, else <c>false</c>.</returns>
    public bool SetDatabaseStatus(string connectionString,
                                  ExternMsgStatus status = ExternMsgStatus.Complete)
    {

      if (string.IsNullOrWhiteSpace(connectionString))
        throw new INExternMsgException("Connection string is required.");

      if (string.IsNullOrWhiteSpace(this.MessageId))
        throw new INExternMsgException(
            "Message ID is required when setting the database status of a message.");

      if (status == ExternMsgStatus.Undefined)
        throw new ArgumentException(string.Format(
            "Message status cannot be set to '{0}'", status), this.MessageId);

      if (this.Direction != ExternMsgDirection.Outbound)
        throw new INExternMsgException(
            string.Format(
            "Direction '{0}' is invalid when setting the database status of a message.",
            Direction), this.MessageId);

      var endTime = DateTime.MinValue;

      if ((status == ExternMsgStatus.Complete ||
          status == ExternMsgStatus.Error) &&
          (this.EndTime ?? INExternMsgHelper.IMAGENOW_DEFAULT_DATE) <= INExternMsgHelper.IMAGENOW_DEFAULT_DATE)
        endTime = DateTime.Now;

      var setStatus = false;

      try
      {
        using (var sqlConn = new SqlConnection(connectionString))
        {
          sqlConn.Open();

          using (var sqlCmd = new SqlCommand() { Connection = sqlConn })
          {
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@MsgId", this.MessageId);
            sqlCmd.Parameters.AddWithValue("@MsgStat", (int)status);

            if (endTime == DateTime.MinValue)
            {
              // If we didn't generate an end time, use the one already 
              // defined in the current instance
              if (!EndTime.HasValue)
                sqlCmd.Parameters.AddWithValue("@EndTime", DBNull.Value);
              else
                sqlCmd.Parameters.AddWithValue("@EndTime", EndTime);
            }
            else
              // Otherwise, use the end time we generated above
              sqlCmd.Parameters.AddWithValue("@EndTime", endTime);

            sqlCmd.CommandText = INExternMsgHelper.IN_EXTERN_MSG_UPDATE_STATUS_CMD;

            var numRows = sqlCmd.ExecuteNonQuery();

            if (numRows > 0)
              setStatus = true;
          }
        }
      }
      catch
      {
        setStatus = false;
      }

      if (setStatus)
      {
        Status = status;

        if (endTime != DateTime.MinValue)
          EndTime = endTime;
      }

      return setStatus;
    }

    /// <summary>
    /// Sends the current <see cref="INExternMsg"/> to the database.
    /// </summary>
    /// <param name="connectionString">The ImageNow database connection string.</param>
    public void Send(string connectionString)
    {
      ValidateMessage();

      // Specify a message ID if needed
      if (string.IsNullOrWhiteSpace(MessageId))
        MessageId = Guid.NewGuid().ToString();

      if (this.Direction == ExternMsgDirection.Undefined)
        throw new INExternMsgException(string.Format(
            "Message direction '{0}' is invalid when sending a message.",
            this.Direction), this.MessageId);

      if (this.Status != ExternMsgStatus.New)
        throw new INExternMsgException(string.Format(
            "Message status must be set to '{0}' when sending a message.",
            ExternMsgStatus.New), this.MessageId);

      // Set the start time if needed
      if (StartTime <= INExternMsgHelper.IMAGENOW_DEFAULT_DATE)
        StartTime = DateTime.Now;

      try
      {
        var inserted = InternalSendMessage(connectionString);

        if (!inserted)
          throw new INExternMsgException("An unknown error occurred while sending the message.", MessageId);

        OnMessageSent();
      }
      catch (Exception e)
      {
        throw new INExternMsgException(
            string.Format(
            "Could not create new {0} object in the database: {1}",
            typeof(INExternMsg).Name, e.Message), MessageId, e);
      }
    }

    /// <summary>
    /// Deletes the current <see cref="INExternMsg"/> from the database.
    /// </summary>
    /// <param name="connectionString">The ImageNow database connection string.</param>
    /// <returns><c>True</c> if the message was successfully deleted from the 
    /// database, else <c>false</c>.</returns>
    public bool Delete(string connectionString)
    {
      var deleted = false;

      if (string.IsNullOrWhiteSpace(connectionString))
        throw new ArgumentException("Connection string is required.", nameof(connectionString));

      if (string.IsNullOrWhiteSpace(MessageId))
        throw new INExternMsgException("Message ID is required when deleting a message.");

      try
      {
        using (var sqlConn = new SqlConnection(connectionString))
        {
          sqlConn.Open();

          using (var sqlCmd = new SqlCommand() { Connection = sqlConn })
          {
            sqlCmd.CommandType = CommandType.Text;
            sqlCmd.Parameters.AddWithValue("@MsgId", MessageId);
            sqlCmd.CommandText = INExternMsgHelper.IN_EXTERN_MSG_DELETE_PROPERTIES_CMD;

            var numRows = sqlCmd.ExecuteNonQuery();

            deleted = numRows >= 0;
          }

          if (deleted)
          {
            using (var sqlCmd = new SqlCommand() { Connection = sqlConn })
            {
              sqlCmd.CommandType = CommandType.Text;
              sqlCmd.Parameters.AddWithValue("@MsgId", MessageId);
              sqlCmd.CommandText = INExternMsgHelper.IN_EXTERN_MSG_DELETE_MSG_CMD;

              var numRows = sqlCmd.ExecuteNonQuery();

              deleted = numRows > 0;
            }
          }
        }
      }
      catch
      {
        deleted = false;
      }

      return deleted;
    }

    /// <summary>
    /// Determines whether the specified object is equal to the current <see cref="INExternMsg"/>.
    /// </summary>
    /// <param name="obj">Object to compare against the current instance.</param>
    /// <returns><c>True</c> on equality, else <c>false</c>.</returns>
    public override bool Equals(object obj)
    {
      if (obj is null)
        return false;

      var inExternMsg = obj as INExternMsg;

      return !(inExternMsg is null) &&
              this.Equals(inExternMsg);
    }

    /// <summary>
    /// Determines whether the specified <see cref="INExternMsg"/> is equal to the 
    /// current <see cref="INExternMsg"/>.
    /// </summary>
    /// <param name="message"><see cref="INExternMsg"/> object to compare against the 
    /// current instance.</param>
    /// <returns><c>True</c> on equality, else <c>false</c>.</returns>
    public bool Equals(INExternMsg message)
    {
      if (message is null)
        return false;

      return Type.Equals(message.Type, StringComparison) &&
          Name.Equals(message.Name, StringComparison) &&
          MessageId.Equals(message.MessageId, StringComparison) &&
          !string.IsNullOrWhiteSpace(MessageId) && !string.IsNullOrWhiteSpace(message.MessageId);
    }

    /// <summary>
    /// Returns the hash code for the current <see cref="INExternMsg"/>.
    /// </summary>
    /// <returns>Hash code for the current <see cref="INExternMsg"/>.</returns>
    public override int GetHashCode()
    {
      // Set initial value
      int hash = 13;

      if (!string.IsNullOrWhiteSpace(MessageId))
      {
        var messageId = MessageId;
        var name = Name;
        var type = Type;

        switch (StringComparison)
        {
          case StringComparison.CurrentCultureIgnoreCase:
            messageId = messageId.ToUpper(CultureInfo.CurrentCulture);
            name = name.ToUpper(CultureInfo.CurrentCulture);
            type = type.ToUpper(CultureInfo.CurrentCulture);
            break;

          case StringComparison.InvariantCultureIgnoreCase:
          case StringComparison.OrdinalIgnoreCase:
            messageId = messageId.ToUpperInvariant();
            name = name.ToUpperInvariant();
            type = type.ToUpperInvariant();
            break;
        }

        hash = (hash * 7) + messageId.GetHashCode();
        hash = (hash * 7) + name.GetHashCode();
        hash = (hash * 7) + type.GetHashCode();
      }
      else
        hash = (hash * 3) + new Random(DateTime.Now.Millisecond +
            new Random(Guid.NewGuid().GetHashCode()).Next()).Next();

      hash = (hash * 3) + this.GetType().GetHashCode();

      // Return the hash code
      return hash;
    }

    /// <summary>
    /// Returns a string representation of the current <see cref="INExternMsg"/>.
    /// </summary>
    /// <returns>String representation of the current <see cref="INExternMsg"/>.</returns>
    public override string ToString()
    {
      return ToString("G", CultureInfo.CurrentCulture);
    }

    /// <summary>
    /// Returns a string representation of the current <see cref="INExternMsg"/>.
    /// </summary>
    /// <param name="format">The format string.</param>
    /// <returns>String representation of the current <see cref="INExternMsg"/>.</returns>
    public string ToString(string format)
    {
      return ToString(format, CultureInfo.CurrentCulture);
    }

    /// <summary>
    /// Returns a string representation of the current <see cref="INExternMsg"/>.
    /// </summary>
    /// <param name="provider">The format provider containing culture-specific 
    /// information relevant for formatting the <see cref="INExternMsg"/>.</param>
    /// <returns>String representation of the current <see cref="INExternMsg"/>.</returns>
    public string ToString(IFormatProvider provider)
    {
      return ToString("G", provider);
    }

    /// <summary>
    /// Returns a string representation of the current <see cref="INExternMsg"/>.
    /// </summary>
    /// <param name="format">The format string.</param>
    /// <param name="provider">The format provider containing culture-specific 
    /// information relevant for formatting the <see cref="INExternMsg"/>.</param>
    /// <returns>String representation of the current <see cref="INExternMsg"/>.</returns>
    public string ToString(string format, IFormatProvider provider)
    {
      var retVal = string.Empty;

      if (string.IsNullOrEmpty(format))
        format = "G";

      if (provider == null)
        provider = CultureInfo.CurrentCulture;

      switch (format)
      {
        // General format
        case "G":
        case "g":
          retVal = string.Format("Type: {0}, Name: {1}, Property Count: {2}",
              ToString("T", provider), ToString("N", provider),
              ToString("C", provider));
          break;

        // Message name
        case "N":
        case "n":
          retVal = Name ?? string.Empty;
          break;

        // Message type
        case "T":
        case "t":
          retVal = Type ?? string.Empty;
          break;

        // End time
        case "E":
        case "e":
          retVal = EndTime.HasValue && EndTime.Value >= INExternMsgHelper.IMAGENOW_DEFAULT_DATE ?
              EndTime.Value.ToString("G", provider) : string.Empty;
          break;

        // Start time
        case "S":
        case "s":
          retVal = StartTime >= INExternMsgHelper.IMAGENOW_DEFAULT_DATE ?
              StartTime.ToString("G", provider) : string.Empty;
          break;

        // Message ID
        case "I":
        case "i":
          retVal = MessageId ?? string.Empty;
          break;

        // Message direction
        case "D":
          retVal = Direction.ToString();
          break;

        // Message status
        case "U":
          retVal = Status.ToString();
          break;

        // Message direction as a numeric string
        case "d":
          retVal = ((int)Direction).ToString(provider);
          break;

        // Message status as a numeric string
        case "u":
          retVal = ((int)Status).ToString(provider);
          break;

        // List of property names (delimited by new line)
        case "P":
        case "p":
          foreach (var p in GetPropertyNames())
            retVal = string.Format("{0}{1}{2}", retVal, p, Environment.NewLine);
          break;

        // List of property names and their values (delimited by new line)
        case "V":
        case "v":
          foreach (var p in GetPropertyNames())
            retVal = string.Format("{0}{1}: {2}{3}", retVal, p,
                this[p] ?? string.Empty, Environment.NewLine);
          break;

        // Number of properties defined
        case "C":
        case "c":
          retVal = PropertyCount.ToString(provider);
          break;

        // String comparison as a string
        case "A":
          retVal = StringComparison.ToString();
          break;

        // String comparison as a numeric string
        case "a":
          retVal = ((int)StringComparison).ToString(provider);
          break;

        // Object type (name only)
        case "o":
          retVal = GetType().Name;
          break;

        // Object type (full)
        case "O":
          retVal = GetType().FullName;
          break;

        // Default, throw exception
        default:
          throw new FormatException(string.Format(
              "Format string '{0}' is not supported.", format));
      }

      return retVal;
    }

    #endregion

    #region Private instance methods

    /// <summary>
    /// Validates the required class properties of the message.
    /// </summary>
    void ValidateMessage()
    {
      if (string.IsNullOrWhiteSpace(Type))
        throw new INExternMsgException("Message type is required.", MessageId);

      if (string.IsNullOrWhiteSpace(Name))
        throw new INExternMsgException("Message name is required.", MessageId);
    }

    /// <summary>
    /// Validates the specified property name and value to ensure that 
    /// the property name is not null/empty/whitespace and that neither the property name nor 
    /// the property value length exceed the ImageNow database size limits.
    /// </summary>
    /// <param name="propertyName">Name of the property.</param>
    /// <param name="propertyVal">Value of the property.</param>
    /// <param name="validatePropVal"><c>True</c> to validate the property name and value, else <c>false</c> to
    /// only validate the property name.</param>
    void ValidatePropNameAndValue(string propertyName,
                                  string propertyVal,
                                  bool validatePropVal = true)
    {
      if (string.IsNullOrWhiteSpace(propertyName))
        throw new ArgumentException("Property name is required.");

      if (propertyName.Length > INExternMsgHelper.MAX_PROP_NAME_LEN)
        throw new ArgumentException(
            string.Format("Property name cannot be more than {0} characters in length.",
            INExternMsgHelper.MAX_PROP_NAME_LEN));

      if (validatePropVal)
      {
        if (!string.IsNullOrEmpty(propertyVal))
        {
          if (propertyVal.Length > INExternMsgHelper.MAX_PROP_VAL_LEN)
            throw new ArgumentException(
                string.Format("Property value cannot be more than {0} characters in length.",
                INExternMsgHelper.MAX_PROP_VAL_LEN));
        }
      }
    }

    /// <summary>
    /// Inserts the message and its properties into the ImageNow database.
    /// </summary>
    /// <param name="connectionString">The ImageNow database connection string.</param>
    /// <returns><c>True</c> if at least one row was inserted into the IN_EXTERN_MSG table and if
    /// the number of rows inserted into the IN_EXTERN_MSG_PROP table matches the number 
    /// of properties defined for the current <see cref="INExternMsg"/>, else <c>false</c>.</returns>
    bool InternalSendMessage(string connectionString)
    {
      var numRowsInserted = 0;
      var numPropRowsInserted = 0;

      using (var sqlConn = new SqlConnection(connectionString))
      {
        sqlConn.Open();

        using (var sqlCmd = new SqlCommand() { Connection = sqlConn })
        {
          sqlCmd.CommandType = CommandType.Text;
          sqlCmd.Parameters.AddWithValue("@MsgId", this.MessageId);
          sqlCmd.Parameters.AddWithValue("@MsgType", this.Type);
          sqlCmd.Parameters.AddWithValue("@MsgName", this.Name);
          sqlCmd.Parameters.AddWithValue("@MsgDir", (int)this.Direction);
          sqlCmd.Parameters.AddWithValue("@MsgStat", (int)this.Status);
          sqlCmd.Parameters.AddWithValue("@StartTime", this.StartTime);

          if (this.EndTime == null)
            sqlCmd.Parameters.AddWithValue("@EndTime", DBNull.Value);
          else
            sqlCmd.Parameters.AddWithValue("@EndTime", this.EndTime);

          sqlCmd.CommandText = INExternMsgHelper.IN_EXTERN_MSG_INSERT_CMD;

          numRowsInserted = sqlCmd.ExecuteNonQuery();
        }

        // Only proceed if at least one row was inserted to the IN_EXTERN_MSG table
        if (numRowsInserted > 0)
        {
          foreach (var p in properties)
          {
            using (var sqlCmd = new SqlCommand() { Connection = sqlConn })
            {
              sqlCmd.CommandType = CommandType.Text;
              sqlCmd.Parameters.AddWithValue("@MsgId", this.MessageId);
              sqlCmd.Parameters.AddWithValue("@PropName", p.Key);
              sqlCmd.Parameters.AddWithValue("@PropType", (int)ExternMsgPropType.Undefined);

              if (p.Value == null)
                sqlCmd.Parameters.AddWithValue("@PropVal", DBNull.Value);
              else
                sqlCmd.Parameters.AddWithValue("@PropVal", p.Value);

              sqlCmd.CommandText = INExternMsgHelper.IN_EXTERN_MSG_PROP_INSERT_CMD;

              numPropRowsInserted += sqlCmd.ExecuteNonQuery();
            }
          }
        }
      }

      return numRowsInserted > 0 && numPropRowsInserted == PropertyCount;
    }

    /// <summary>
    /// Determines whether the specified property name is defined in the current 
    /// <see cref="INExternMsg"/> by performing search of all defined property names in 
    /// the current instance.
    /// </summary>
    /// <param name="propertyName">Name of the property to find.</param>
    /// <param name="definedPropertyName">The actual property name as defined in the 
    /// internal dictionary of property names and values.</param>
    /// <returns><c>True</c> if the property name exists, else <c>false</c>.</returns>
    bool InternalContains(string propertyName, out string definedPropertyName)
    {
      propertyName = propertyName ?? string.Empty;
      definedPropertyName = null;
      var contains = false;

      if (properties != null)
      {
        definedPropertyName = properties.Keys.Where(p => !string.IsNullOrWhiteSpace(p) &&
                                                         p.Equals(propertyName, StringComparison))
                                        .FirstOrDefault();

        contains = !string.IsNullOrEmpty(definedPropertyName);
      }

      return contains;
    }

    #endregion

    #region Protected instance methods

    /// <summary>
    /// Raises the <see cref="MessageSent"/> event on the current <see cref="INExternMsg"/>.
    /// </summary>
    protected void OnMessageSent()
    {
      MessageSent?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Raises the <see cref="PropertyValueChanged"/> event on the current <see cref="INExternMsg"/>.
    /// </summary>
    /// <param name="propertyName">The name of the property being changed.</param>
    /// <param name="oldValue">The value of the property before being changed.</param>
    /// <param name="newValue">The value of the property after being changed.</param>
    protected void OnPropertyValueChanged(string propertyName,
                                          string oldValue,
                                          string newValue)
    {
      if (PropertyValueChanged != null)
      {

        var e = new PropertyValueChangedEventArgs()
        {
          PropertyName = propertyName,
          OldPropertyValue = oldValue,
          NewPropertyValue = newValue,
        };

        PropertyValueChanged(this, e);
      }
    }

    /// <summary>
    /// Raises the <see cref="PropertyAdded"/> event on the current <see cref="INExternMsg"/>.
    /// </summary>
    /// <param name="propertyName">The name of the property being added.</param>
    /// <param name="propertyValue">The value of the property being added.</param>
    protected void OnPropertyAdded(string propertyName, string propertyValue)
    {
      if (PropertyAdded != null)
      {

        var e = new PropertyAddedEventArgs()
        {
          PropertyName = propertyName,
          PropertyValue = propertyValue,
        };

        PropertyAdded(this, e);
      }
    }

    #endregion

    #region Static methods

    /// <summary>
    /// Retrieves a single <see cref="INExternMsg"/> using the specified connection string and 
    /// message ID.
    /// </summary>
    /// <param name="connectionString">The ImageNow database connection string.</param>
    /// <param name="messageId">The ID of the message that will be retrieved.</param>
    /// <returns><see cref="INExternMsg"/> with the specified ID.</returns>
    public static INExternMsg FromDatabase(string connectionString,
                                           string messageId)
    {

      if (string.IsNullOrWhiteSpace(connectionString))
        throw new ArgumentException("Connection string is required.", nameof(connectionString));

      if (string.IsNullOrWhiteSpace(messageId))
        throw new ArgumentException("Message ID is required.", nameof(messageId));

      return new INExternMsgReader(connectionString).Get(messageId);
    }

    /// <summary>
    /// Creates <see cref="INExternMsg"/> objects using the contents of the specified delimited file.
    /// </summary>
    /// <remarks>The delimited file should be in the following format, assuming the delimiter is a comma:
    /// 
    /// Message ID,Message Type,Message Name,Property 1 Name,Property 1 Value,Property 2 Name,Property 2 Value
    /// 
    /// The number of fields on each line (or message) in the file may vary, depending on how many 
    /// properties are defined for each message.
    /// 
    /// The message ID field may be left blank; the value will be auto-generated when the message is sent.
    /// 
    /// The first three columns of the file are required. If no value is needed for one of the three required columns, 
    /// use an empty string as the value.
    /// 
    /// Example:
    /// 
    /// ,AP,Invoice,Amount,156.78,Vendor,"Acme Consulting, Inc.",Date,5/27/2018
    /// 
    /// This line defines a new message of type 'AP' and name 'Invoice' with no message ID specified. Three properties 
    /// are defined with values: 'Amount' (156.78), 'Vendor' (Acme Consulting, Inc.), and 'Date' (5/27/2018).
    /// </remarks>
    /// <param name="filePath">The path to the delimited file.</param>
    /// <param name="encoding">The <see cref="Encoding"/> applied to the contents of the file.</param>
    /// <param name="numberOfHeaderRows">The number of header rows present at the beginning of the file.</param>
    /// <param name="delimiter">The field delimiter in the file.</param>
    /// <param name="quotedBy">The optional character used for quoting the delimiter or other special characters.</param>
    /// <returns>Array of <see cref="INExternMsg"/> objects read from the contents of the delimited file.</returns>
    public static INExternMsg[] FromFile(string filePath,
                                         Encoding encoding,
                                         int numberOfHeaderRows = 0,
                                         char delimiter = ',',
                                         char quotedBy = '"')
    {

      if (string.IsNullOrWhiteSpace(filePath))
        throw new ArgumentException("File path is required.", nameof(filePath));

      var fileInfo = new FileInfo(filePath);

      if (!fileInfo.Exists)
        throw new FileNotFoundException($"File '{filePath}' does not exist.");

      if (encoding == null)
        encoding = Encoding.ASCII;

      var contents = File.ReadAllLines(filePath, encoding);

      if (contents.Length < 1)
        return null;

      var messages = new List<INExternMsg>();

      var numHeaderRowsSkipped = 0;

      foreach (var line in contents)
      {
        if (numberOfHeaderRows > 0 && numHeaderRowsSkipped != numberOfHeaderRows)
        {
          numHeaderRowsSkipped++;
          continue;
        }

        if (string.IsNullOrEmpty(line))
          continue;

        var fields = ArrayFromDelimitedString(line, delimiter, quotedBy);

        if (fields.Length < 3)
          throw new InvalidDataException($"Malformed line: {line}");

        var msg = new INExternMsg()
        {
          MessageId = fields[0],
          Name = fields[2],
          Type = fields[1],
          Status = ExternMsgStatus.New,
          Direction = ExternMsgDirection.Inbound,
          StartTime = DateTime.Now
        };

        if (fields.Length > 3)
        {
          for (var x = 3; x < fields.Length; x++)
          {
            string propName = fields[x];
            string propVal = null;

            if (fields.Length > x + 1)
              propVal = fields[++x];

            msg.SetProperty(propName, propVal, ExternMsgPropType.Undefined);
          }
        }

        messages.Add(msg);
      }
      return messages.ToArray();
    }

    /// <summary>
    /// Creates a new <see cref="INExternMsg"/> using the specified <see cref="IDictionary{string, object}"/> 
    /// as the source.
    /// </summary>
    /// <param name="dictionary">Source dictionary.</param>
    /// <returns><see cref="INExternMsg"/> from the source <see cref="IDictionary{string, object}"/>.</returns>
    public static INExternMsg FromDictionary(IDictionary<string, object> dictionary)
    {
      var message = new INExternMsg();

      if (dictionary != null)
      {
        foreach (var keyValuePair in dictionary)
        {
          if (string.IsNullOrEmpty(keyValuePair.Key))
            continue;

          switch (keyValuePair.Key.ToUpperInvariant())
          {
            case "MSGID":
              message.MessageId = keyValuePair.Value?.ToString();
              break;

            case "MSGNAME":
              message.Name = keyValuePair.Value?.ToString();
              break;

            case "MSGTYPE":
              message.Type = keyValuePair.Value?.ToString();
              break;

            case "MSGSTART":
              if (keyValuePair.Value is DateTime dateTimeStart)
                message.StartTime = dateTimeStart;
              else if (DateTime.TryParse(keyValuePair.Value?.ToString(), out DateTime parsedDateTimeStart))
                message.StartTime = parsedDateTimeStart;
              break;

            case "MSGEND":
              if (keyValuePair.Value is DateTime dateTimeEnd)
                message.EndTime = dateTimeEnd;
              else if (DateTime.TryParse(keyValuePair.Value?.ToString(), out DateTime parsedDateTimeEnd))
                message.EndTime = parsedDateTimeEnd;
              break;

            case "MSGDIRECTION":
            case "MSGDIR":
              if (keyValuePair.Value is ExternMsgDirection direction)
                message.Direction = direction;
              else if (Enum.TryParse(keyValuePair.Value?.ToString(), true, out ExternMsgDirection parsedDirection))
                message.Direction = parsedDirection;
              break;

            case "MSGSTATUS":
              if (keyValuePair.Value is ExternMsgStatus status)
                message.Status = status;
              else if (Enum.TryParse(keyValuePair.Value?.ToString(), true, out ExternMsgStatus parsedStatus))
                message.Status = parsedStatus;
              break;

            default:
              message.SetProperty(keyValuePair.Key, keyValuePair.Value?.ToString(), ExternMsgPropType.Undefined);
              break;
          }
        }
      }

      return message;
    }

    /// <summary>
    /// Creates an array of <see cref="INExternMsg"/> objects from the specified array of 
    /// <see cref="IDictionary{string, object}"/> objects.
    /// </summary>
    /// <param name="dictionaries">Source dictionaries.</param>
    /// <returns>Array of <see cref="INExternMsg"/> objects from the specified array of 
    /// <see cref="IDictionary{string, object}"/> objects</returns>
    public static INExternMsg[] FromDictionary(params IDictionary<string, object>[] dictionaries)
    {
      var messageList = new List<INExternMsg>();

      if (dictionaries != null)
      {
        foreach (var dictionary in dictionaries)
        {
          var message = FromDictionary(dictionary);

          if (message != null)
            messageList.Add(message);
        }
      }

      return messageList.ToArray();
    }

    /// <summary>
    /// Sends the specified <see cref="INExternMsg"/> objects to the database for processing.
    /// </summary>
    /// <param name="connectionString">The ImageNow database connection string.</param>
    /// <param name="messages"><see cref="INExternMsg"/> objects that will be sent.</param>
    public static void Send(string connectionString,
                            params INExternMsg[] messages)
    {

      if (string.IsNullOrWhiteSpace(connectionString))
        throw new ArgumentException("Connection string is required.", nameof(connectionString));

      if (messages?.Length > 0)
      {
        var exceptions = new List<Exception>();

        foreach (var message in messages)
        {
          try
          {
            message?.Send(connectionString);
          }
          catch (Exception ex)
          {
            exceptions.Add(ex);
            continue;
          }
        }

        if (exceptions.Count > 0)
          throw new AggregateException($"{(exceptions.Count == 1 ? "An error" : $"Errors ({exceptions.Count})")} occurred while sending messages.",
                                      exceptions.ToArray());
      }
    }

    /// <summary>
    /// Splits the specified <paramref name="source"/> text delimited by the specified <paramref name="delimiter"/>.
    /// </summary>
    /// <param name="source">Source text.</param>
    /// <param name="delimiter">Delimiter character</param>
    /// <param name="quotedBy">Quote character.</param>
    /// <returns>String array of values from the source text.</returns>
    static string[] ArrayFromDelimitedString(string source, char delimiter, char quotedBy)
    {
      string[] values;

      var doubleQuoteStart = source.IndexOf(quotedBy.ToString(), StringComparison.Ordinal);
      if (doubleQuoteStart > -1)
      {
        var startOfData = 0;
        var tempData = new List<string>();

        while (startOfData < source.Length)
        {
          var iNextComma = source.IndexOf(delimiter.ToString(), startOfData, StringComparison.Ordinal);
          if (iNextComma == -1)
            iNextComma = source.Length;
          if (iNextComma < doubleQuoteStart || doubleQuoteStart == -1)
          {
            tempData.Add(iNextComma - startOfData > 0 ? source.Substring(startOfData, iNextComma - startOfData) : "");
            startOfData = iNextComma + 1;
          }
          else
          {
            var doubleQuoteEnd = source.IndexOf(quotedBy.ToString(), doubleQuoteStart + 1, StringComparison.Ordinal);
            tempData.Add(source.Substring(doubleQuoteStart + 1, doubleQuoteEnd - doubleQuoteStart - 1));
            startOfData = (doubleQuoteEnd + 2) > source.Length ? doubleQuoteEnd + 1 : doubleQuoteEnd + 2;
            doubleQuoteStart = source.IndexOf(quotedBy.ToString(), startOfData, StringComparison.Ordinal);
          }
        }

        values = tempData.ToArray();
      }
      else
      {
        values = source.Split(delimiter);
      }

      return values;
    }

    #endregion

    #region Private variables

    /// <summary>
    /// Internal <see cref="Dictionary{TKey, TValue}"/> object that holds the message properties.
    /// </summary>
    Dictionary<string, string> properties;

    /// <summary>
    /// The message end time.
    /// </summary>
    DateTime? endTime;

    /// <summary>
    /// The message start time.
    /// </summary>
    DateTime? startTime;

    /// <summary>
    /// The message type, name, and ID.
    /// </summary>
    string msgType, msgName, msgId;

    #endregion
  }
}
