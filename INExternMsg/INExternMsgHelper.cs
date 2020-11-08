using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace org.goodspace.Utils.ImageNow {

  static class INExternMsgHelper {

    #region Constants

    /// <summary>
    /// The default ImageNow database date/time value
    /// </summary>
    public static DateTime IMAGENOW_DEFAULT_DATE = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);

    /// <summary>
    /// SQL command for inserting into the IN_EXTERN_MSG table in ImageNow.
    /// </summary>
    public const string IN_EXTERN_MSG_INSERT_CMD = "INSERT INTO " + IN_EXTERN_MSG_TABLE +
        " (" + MESSAGE_ID_COL + ", " + TYPE_COL + ", " + NAME_COL + ", " + DIRECTION_COL +
        ", " + STATUS_COL + ", " + START_TIME_COL + ", " + END_TIME_COL +
        ") VALUES (@MsgId, @MsgType, @MsgName, @MsgDir, " +
        "@MsgStat, @StartTime, @EndTime)";

    /// <summary>
    /// SQL command for inserting into the IN_EXTERN_MSG_PROP table in ImageNow.
    /// </summary>
    public const string IN_EXTERN_MSG_PROP_INSERT_CMD = "INSERT INTO " +
        IN_EXTERN_MSG_PROP_TABLE + " (" + MESSAGE_ID_COL + ", " + PROP_NAME_COL + ", " +
        PROP_TYPE_COL + ", " + PROP_VAL_COL + ") VALUES (" +
        "@MsgId, @PropName, @PropType, @PropVal)";

    /// <summary>
    /// SQL command for retrieving new, outbound messages of a specific type with a 
    /// specific name from the IN_EXTERN_MSG table in ImageNow.
    /// </summary>
    public const string IN_EXTERN_MSG_SELECT_CMD = "SELECT " + MESSAGE_ID_COL + ", " +
        TYPE_COL + ", " + NAME_COL + ", " + DIRECTION_COL + ", " + STATUS_COL +
        ", " + START_TIME_COL + ", " + END_TIME_COL + " FROM " + IN_EXTERN_MSG_TABLE +
        " WHERE " + NAME_COL + " = @MsgName AND " + TYPE_COL + " = @MsgType AND " +
        DIRECTION_COL + " = 2 AND " + STATUS_COL + " = 1 ORDER BY " + START_TIME_COL;

    /// <summary>
    /// SQL command for retrieving all new, outbound messages from 
    /// the IN_EXTERN_MSG table in ImageNow.
    /// </summary>
    public const string IN_EXTERN_MSG_SELECT_ALL_CMD = "SELECT " + MESSAGE_ID_COL + ", " +
        TYPE_COL + ", " + NAME_COL + ", " + DIRECTION_COL + ", " + STATUS_COL +
        ", " + START_TIME_COL + ", " + END_TIME_COL + " FROM " + IN_EXTERN_MSG_TABLE +
        " WHERE " + DIRECTION_COL + " = 2 AND " + STATUS_COL + " = 1 ORDER BY " +
        START_TIME_COL;

    /// <summary>
    /// SQL command for retrieving a single message using 
    /// the message ID from the IN_EXTERN_MSG table in ImageNow.
    /// </summary>
    public const string IN_EXTERN_MSG_SELECT_ID_CMD = "SELECT " + MESSAGE_ID_COL +
        ", " + TYPE_COL + ", " + NAME_COL + ", " + DIRECTION_COL + ", " + STATUS_COL +
        ", " + START_TIME_COL + ", " + END_TIME_COL + " FROM " + IN_EXTERN_MSG_TABLE +
        " WHERE " + MESSAGE_ID_COL + " = @MsgId";

    /// <summary>
    /// SQL command for retrieving new, outbound messages of a certain 
    /// type from the IN_EXTERN_MSG table in ImageNow.
    /// </summary>
    public const string IN_EXTERN_MSG_SELECT_TYPE_CMD = "SELECT " + MESSAGE_ID_COL + ", "
        + TYPE_COL + ", " + NAME_COL + ", " + DIRECTION_COL + ", " + STATUS_COL +
        ", " + START_TIME_COL + ", " + END_TIME_COL + " FROM " + IN_EXTERN_MSG_TABLE +
        " WHERE " + TYPE_COL + " = @MsgType AND " + DIRECTION_COL +
        " = 2 AND " + STATUS_COL + " = 1 ORDER BY " + START_TIME_COL;

    /// <summary>
    /// SQL command for retrieving the properties of a message from the 
    /// IN_EXTERN_MSG_PROP table in ImageNow.
    /// </summary>
    public const string IN_EXTERN_MSG_PROP_SELECT_CMD = "SELECT " + MESSAGE_ID_COL +
        ", " + PROP_NAME_COL + ", " + PROP_TYPE_COL + ", " + PROP_VAL_COL +
        " FROM " + IN_EXTERN_MSG_PROP_TABLE + " WHERE " + MESSAGE_ID_COL + " = @MsgId";

    /// <summary>
    /// SQL command for updating the status of an outbound message 
    /// in the IN_EXTERN_MSG table in ImageNow.
    /// </summary>
    public const string IN_EXTERN_MSG_UPDATE_STATUS_CMD =
        "UPDATE " + IN_EXTERN_MSG_TABLE + " SET " + STATUS_COL + " = @MsgStat, " +
        END_TIME_COL + " = @EndTime WHERE " + MESSAGE_ID_COL + " = @MsgId AND " +
        DIRECTION_COL + " = 2";

    /// <summary>
    /// SQL command for deleting a message from  
    /// the IN_EXTERN_MSG table in ImageNow.
    /// </summary>
    public const string IN_EXTERN_MSG_DELETE_MSG_CMD =
        "DELETE FROM " + IN_EXTERN_MSG_TABLE + " WHERE " + MESSAGE_ID_COL + " = @MsgId";

    /// <summary>
    /// SQL command for deleting message properties from 
    /// the IN_EXTERN_MSG_PROP table in ImageNow.
    /// </summary>
    public const string IN_EXTERN_MSG_DELETE_PROPERTIES_CMD =
        "DELETE FROM " + IN_EXTERN_MSG_PROP_TABLE + " WHERE " + MESSAGE_ID_COL + " = @MsgId";

    /// <summary>
    /// Name of the property name column in the ImageNow database.
    /// </summary>
    public const string PROP_NAME_COL = "PROP_NAME";

    /// <summary>
    /// Name of the property value column in the ImageNow database.
    /// </summary>
    public const string PROP_VAL_COL = "PROP_VALUE";

    /// <summary>
    /// Name of the message ID column in the ImageNow database.
    /// </summary>
    public const string MESSAGE_ID_COL = "EXTERN_MSG_ID";

    /// <summary>
    /// Name of the start time column in the ImageNow database.
    /// </summary>
    public const string START_TIME_COL = "START_TIME";

    /// <summary>
    /// Name of the end time column in the ImageNow database.
    /// </summary>
    public const string END_TIME_COL = "END_TIME";

    /// <summary>
    /// Name of the property type column in the ImageNow database.
    /// </summary>
    public const string PROP_TYPE_COL = "PROP_TYPE";

    /// <summary>
    /// Name of the message name column in the ImageNow database.
    /// </summary>
    public const string NAME_COL = "MSG_NAME";

    /// <summary>
    /// Name of the message type column in the ImageNow database.
    /// </summary>
    public const string TYPE_COL = "MSG_TYPE";

    /// <summary>
    /// Name of the message status column in the ImageNow database.
    /// </summary>
    public const string STATUS_COL = "MSG_STATUS";

    /// <summary>
    /// Name of the external message table in the ImageNow database.
    /// </summary>
    public const string IN_EXTERN_MSG_TABLE = "tmp.IN_EXTERN_MSG";

    /// <summary>
    /// Name of the external message property table in the ImageNow database.
    /// </summary>
    public const string IN_EXTERN_MSG_PROP_TABLE = "tmp.IN_EXTERN_MSG_PROP";

    /// <summary>
    /// Name of the message direction column in the ImageNow database.
    /// </summary>
    public const string DIRECTION_COL = "MSG_DIRECTION";

    /// <summary>
    /// The maximum character length for the message ID column 
    /// in the ImageNow database.
    /// </summary>
    public const int MAX_MESSAGE_ID_LEN = 64;

    /// <summary>
    /// The maximum character length for the property value column 
    /// in the ImageNow database.
    /// </summary>
    public const int MAX_PROP_VAL_LEN = 256;

    /// <summary>
    /// The maximum character length for the property name column 
    /// in the ImageNow database.
    /// </summary>
    public const int MAX_PROP_NAME_LEN = 64;

    /// <summary>
    /// The maximum character length for the message name column 
    /// in the ImageNow database.
    /// </summary>
    public const int MAX_MSG_NAME_LEN = 64;

    /// <summary>
    /// The maximum character length for the message type column 
    /// in the ImageNow database.
    /// </summary>
    public const int MAX_MSG_TYPE_LEN = 64;

    #endregion

    #region Static methods

    public static IEnumerable<INExternMsg> GetMessages(string connectionString,
                                                       string query,
                                                       Dictionary<string, object> parameters,
                                                       StringComparison stringComparison = StringComparison.InvariantCultureIgnoreCase)
    {

      using (var sqlConnection = new SqlConnection(connectionString))
      {

        sqlConnection.Open();

        using (var sqlCmd = new SqlCommand()
        {
          CommandType = CommandType.Text,
          Connection = sqlConnection,
          CommandText = query
        })
        {

          if (parameters != null)
            foreach (var parameter in parameters)
              sqlCmd.Parameters.AddWithValue(parameter.Key, parameter.Value);

          using (var reader = sqlCmd.ExecuteReader())
          {

            while (reader != null && reader.Read())
            {

              var message = BuildMessage(reader, stringComparison);

              var properties = GetProperties(connectionString, message.MessageId);

              if (properties != null)
                foreach (var prop in properties)
                  message.AddProperty(prop.Key, prop.Value, ExternMsgPropType.Undefined);

              yield return message;
            }
          }
        }
      }
    }

    public static INExternMsg BuildMessage(SqlDataReader sqlReader,
                                           StringComparison stringComparison = StringComparison.InvariantCultureIgnoreCase)
    {

      var message = default(INExternMsg);

      if (sqlReader is null || sqlReader.IsClosed || !sqlReader.HasRows)
        return message;

      try
      {

        var msgStat = sqlReader[STATUS_COL] is DBNull ? 0 : (int)sqlReader[STATUS_COL];
        var msgDir = sqlReader[DIRECTION_COL] is DBNull ? 0 : (int)sqlReader[DIRECTION_COL];

        if (!Enum.IsDefined(typeof(ExternMsgDirection), msgDir))
          throw new INExternMsgException(string.Format(
              "Message direction '{0}' is not valid.", msgDir));

        if (!Enum.IsDefined(typeof(ExternMsgStatus), msgStat))
          throw new INExternMsgException(string.Format(
              "Message status '{0}' is not valid.", msgStat));

        message = new INExternMsg(sqlReader[TYPE_COL] is DBNull ? string.Empty :
            (string)sqlReader[TYPE_COL], sqlReader[NAME_COL] is DBNull ? string.Empty :
            (string)sqlReader[NAME_COL], (ExternMsgDirection)msgDir,
            (ExternMsgStatus)msgStat, sqlReader[START_TIME_COL] is DBNull ? IMAGENOW_DEFAULT_DATE :
            (DateTime)sqlReader[START_TIME_COL], sqlReader[END_TIME_COL] is DBNull ? (DateTime?)null :
            (DateTime)sqlReader[END_TIME_COL], stringComparison)
        {
          MessageId = sqlReader[MESSAGE_ID_COL] is DBNull ? null :
                            (string)sqlReader[MESSAGE_ID_COL],
        };

        if (string.IsNullOrWhiteSpace(message.MessageId))
          throw new INExternMsgException("Invalid message ID encountered.");
      }
      catch (Exception ex)
      {
        throw new INExternMsgException(
            string.Format("Error while reading message contents: {0}", ex.Message), ex);
      }

      return message;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="connectionString"></param>
    /// <param name="messageId"></param>
    /// <returns></returns>
    public static Dictionary<string, string> GetProperties(string connectionString,
                                                           string messageId)
    {

      var properties = new Dictionary<string, string>();

      using (var sqlConnection = new SqlConnection(connectionString))
      {

        sqlConnection.Open();

        using (var propertyCmd = new SqlCommand()
        {
          Connection = sqlConnection,
          CommandType = CommandType.Text,
          CommandText = IN_EXTERN_MSG_PROP_SELECT_CMD
        })
        {

          propertyCmd.Parameters.AddWithValue("@MsgId", messageId);

          using (var propertyReader = propertyCmd.ExecuteReader())
          {

            while (propertyReader != null && propertyReader.Read())
            {

              var propName = propertyReader[PROP_NAME_COL] is DBNull ? string.Empty :
                  (string)propertyReader[PROP_NAME_COL];

              var propVal = propertyReader[PROP_VAL_COL] is DBNull ? null :
                  (string)propertyReader[PROP_VAL_COL];

              if (string.IsNullOrWhiteSpace(propName))
                throw new INExternMsgException("Invalid property name encountered.");

              if (!properties.ContainsKey(propName))
                properties.Add(propName, propVal);
              else
                properties[propName] = propVal;
            }
          }
        }
      }

      return properties;
    }

    #endregion
  }
}