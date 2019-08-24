Imports System.Net
Imports System.Web.Script.Serialization
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports System.Text

Public Class hspi_netatmo_api

  Private _access_token As String = String.Empty
  Private _refresh_token As String = String.Empty
  Private _scope As String = String.Empty
  Private _expires_in As Integer = 0
  Private _refreshed As New Stopwatch

  Private _querySuccess As ULong = 0
  Private _queryFailure As ULong = 0

  Public Sub New()

  End Sub

  Public Function QuerySuccessCount() As ULong
    Return _querySuccess
  End Function

  Public Function QueryFailureCount() As ULong
    Return _queryFailure
  End Function

#Region "Netatmo Authentiation"

  ''' <summary>
  ''' Determines if the API key is valid
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ValidAPIToken() As Boolean

    Try

      If gAPIClientId.Length = 0 Then
        Return False
      ElseIf gAPIClientSecret.Length = 0 Then
        Return False
      ElseIf NetatmoAPI.GetAccessToken.Length = 0 Then
        Return False
      Else
        Return True
      End If

    Catch pEx As Exception
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Get Access Token
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetAccessToken() As String

    Dim expires_in As Long = _refreshed.ElapsedMilliseconds / 1000

    If CheckCredentials() = False Then
      WriteMessage("Invalid API authentication information.  Please check plug-in options.", MessageType.Error)
    ElseIf _access_token.Length = 0 Then
      GetToken()
    ElseIf expires_in > _expires_in Then
      _access_token = String.Empty
      RefreshAccessToken()
    End If

    Return _access_token

  End Function

  ''' <summary>
  ''' Gets Accesss Token
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub GetToken()

    Try

      Dim data As Byte() = New ASCIIEncoding().GetBytes(String.Format("grant_type={0}&client_id={1}&client_secret={2}&username={3}&password={4}&scope={5}", "password", gAPIClientId, gAPIClientSecret, gAPIUsername, gAPIPassword, gAPIScope))

      Dim strURL As String = String.Format("https://api.netatmo.com/oauth2/token")
      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)

      HTTPWebRequest.Timeout = 1000 * 60
      HTTPWebRequest.Method = "POST"
      HTTPWebRequest.ContentType = "application/x-www-form-urlencoded;charset=UTF-8"
      HTTPWebRequest.ContentLength = data.Length

      Dim myStream As Stream = HTTPWebRequest.GetRequestStream
      myStream.Write(data, 0, data.Length)

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())

          Dim JSONString As String = reader.ReadToEnd()
          ' {"error":"invalid_client"}

          Dim js As New JavaScriptSerializer()
          Dim OAuth20 As oauth_token = js.Deserialize(Of oauth_token)(JSONString)

          If OAuth20.error.Length > 0 Then
            Throw New Exception(OAuth20.error)
          End If

          ' {"access_token":"53d5a8531977598b59a1cd7c|c5b15ee74615b2cfa4d0598611a8f984","refresh_token":"53d5a8531977598b59a1cd7c|7216e7448890d3523ed2dad3ff5a03e5","scope":["read_station"],"expires_in":10800}

          _access_token = OAuth20.access_token
          _refresh_token = OAuth20.refresh_token
          _expires_in = OAuth20.expires_in

          _refreshed.Start()

        End Using

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Debug)

      _access_token = String.Empty
      _refresh_token = String.Empty
      _expires_in = 0

    End Try

  End Sub

  ''' <summary>
  ''' Checks to see if required credentials are available
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function CheckCredentials() As Boolean

    Try

      Dim sbWarning As New StringBuilder

      If gAPIClientId.Length = 0 Then
        sbWarning.Append("Netatmo Client Id")
      End If
      If gAPIClientSecret.Length = 0 Then
        sbWarning.Append("Netatmo Client Secret")
      End If
      If gAPIUsername.Length = 0 Then
        sbWarning.Append("Netatmo Client Username")
      End If
      If gAPIPassword.Length = 0 Then
        sbWarning.Append("Netatmo Client Password")
      End If
      If sbWarning.Length = 0 Then Return True

    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Debug)
    End Try

    Return False

  End Function

  ''' <summary>
  ''' Refreshes the Access Token
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub RefreshAccessToken()

    Try

      Dim data As Byte() = New ASCIIEncoding().GetBytes(String.Format("grant_type={0}&client_id={1}&client_secret={2}&refresh_token={3}", "refresh_token", gAPIClientId, gAPIClientSecret, _refresh_token))

      Dim strURL As String = String.Format("https://api.netatmo.com/oauth2/token")
      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)

      HTTPWebRequest.Timeout = 1000 * 60
      HTTPWebRequest.Method = "POST"
      HTTPWebRequest.ContentType = "application/x-www-form-urlencoded;charset=UTF-8"
      HTTPWebRequest.ContentLength = data.Length

      Dim myStream As Stream = HTTPWebRequest.GetRequestStream
      myStream.Write(data, 0, data.Length)

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())

          Dim JSONString As String = reader.ReadToEnd()

          Dim js As New JavaScriptSerializer()
          Dim OAuth20 As refresh_token = js.Deserialize(Of refresh_token)(JSONString)

          ' {"access_token":"53d5a8531977598b59a1cd7c|c5b15ee74615b2cfa4d0598611a8f984","refresh_token":"53d5a8531977598b59a1cd7c|7216e7448890d3523ed2dad3ff5a03e5","scope":["read_station"],"expires_in":10800}

          _access_token = OAuth20.access_token
          _refresh_token = OAuth20.refresh_token
          _expires_in = OAuth20.expires_in

          _refreshed.Start()

        End Using

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Debug)

      _access_token = String.Empty
      _refresh_token = String.Empty
      _expires_in = 0

    End Try

  End Sub
#End Region

  ''' <summary>
  ''' Gets the Realtime Weather from UltraNetatmo
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetStationData() As StationData

    Dim StationData As New StationData

    Try

      Dim access_token As String = Me.GetAccessToken()
      If access_token.Length = 0 Then
        Throw New Exception("Unable to get Netatmo Access Token.")
      End If

      Dim strURL As String = String.Format("https://api.netatmo.com/api/getstationsdata?access_token={0}&get_favorites={1}", access_token, gAPIGetFavorites.ToString)
      WriteMessage(strURL, MessageType.Debug)

      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)
      HTTPWebRequest.Timeout = 1000 * 60

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())
          Dim JSONString As String = reader.ReadToEnd()

          Dim js As New JavaScriptSerializer()
          StationData = js.Deserialize(Of StationData)(JSONString)

        End Using

      End Using

      _querySuccess += 1
    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Warning)

      _queryFailure += 1

      _expires_in = 0
    End Try

    Return StationData

  End Function

  ''' <summary>
  ''' Gets the Realtime Weather from UltraNetatmo
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetHomeCoachData() As HomeCoachData

    Dim HomeCoachData As New HomeCoachData

    Try

      Dim access_token As String = Me.GetAccessToken()
      If access_token.Length = 0 Then
        Throw New Exception("Unable to get Netatmo Access Token.")
      End If

      Dim strURL As String = String.Format("https://api.netatmo.com/api/gethomecoachsdata?access_token={0}", access_token)
      WriteMessage(strURL, MessageType.Debug)

      Dim HTTPWebRequest As HttpWebRequest = System.Net.WebRequest.Create(strURL)
      HTTPWebRequest.Timeout = 1000 * 60

      Using response As HttpWebResponse = DirectCast(HTTPWebRequest.GetResponse(), HttpWebResponse)

        Using reader = New StreamReader(response.GetResponseStream())
          Dim JSONString As String = reader.ReadToEnd()

          Dim js As New JavaScriptSerializer()
          HomeCoachData = js.Deserialize(Of HomeCoachData)(JSONString)

        End Using

      End Using

      _querySuccess += 1
    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage(pEx.Message, MessageType.Warning)

      _queryFailure += 1

      _expires_in = 0
    End Try

    Return HomeCoachData

  End Function

#Region "UltraNetatmo oAuth Token"

  '{
  '    "access_token": "53d5a8531977598b59a1cd7c|c5b15ee74615b2cfa4d0598611a8f984",
  '    "refresh_token": "53d5a8531977598b59a1cd7c|7216e7448890d3523ed2dad3ff5a03e5",
  '    "expires_in": 10800
  '}

  <Serializable()>
  Private Class oauth_token
    Public Property [error] As String = String.Empty
    Public Property access_token As String = String.Empty   ' Access token for your user
    Public Property refresh_token As String = String.Empty  ' Use this token to get a new access_token once it has expired
    Public Property expires_in As Integer = 0               ' Validity timelaps in seconds
  End Class

  <Serializable()>
  Private Class refresh_token
    Public Property access_token As String = String.Empty
    Public Property refresh_token As String = String.Empty  ' Refresh tokens do not change
    Public Property expires_in As Integer = 0
  End Class

#End Region

#Region "StationData"

  <Serializable()>
  Public Class StationData
    Public Property body As DevicesAndUser
    Public Property status As String = String.Empty
    Public Property time_exec As Double
    Public Property time_server As Double
  End Class

  <Serializable()>
  Public Class DevicesAndUser
    Public Property devices As List(Of Device)
    Public Property user As User
  End Class

  <Serializable()>
  Public Class Device
    Public Property _id As String = String.Empty
    Public Property cipher_id As String = String.Empty
    Public Property last_status_store As Integer = 0
    Public Property modules As List(Of DeviceModule)          ' List of modules associated with the station and their details
    Public Property place As Place
    Public Property station_name As String = String.Empty
    Public Property type As String = String.Empty
    Public Property dashboard_data As DeviceDashboardData     ' Last data measured per device
    Public Property data_type As String()                     ' Array of data measured by the device (e.g. "Temperature","Humidity")
    Public Property co2_calibrating As Boolean = False
    Public Property date_setup As Long = 0                    ' Date when the Weather station was set up
    Public Property module_name As String = String.Empty
    Public Property firmware As Long = 0                      ' Version of the software
    Public Property last_upgrade As Long = 0
    Public Property wifi_status As Integer = 0
  End Class

  <Serializable()>
  Public Class DeviceModule
    Public Property _id As String = String.Empty
    Public Property type As String = String.Empty
    Public Property last_message As Long = 0
    Public Property last_seen As Long = 0
    Public Property dashboard_data As ModuleDashboardData
    Public Property data_type As String()
    Public Property module_name As String = String.Empty
    Public Property last_setup As Long = 0
    Public Property battery_vp As Integer = 0                 ' Current battery status per module.
    Public Property battery_percent As Integer = 0
    Public Property rf_status As Integer = 0                  ' Wifi status per Base station. (86=bad, 56=good) Find more details on the Weather Station page.
    Public Property firmware As Long = 0
  End Class

  <Serializable()>
  Public Class Place
    Public Property altitude As Double = 0
    Public Property city As String = String.Empty
    Public Property country As String = String.Empty
    Public Property timezone As String = String.Empty
    Public Property location As Double()
  End Class

  <Serializable()>
  Public Class ModuleDashboardData
    Public Property time_utc As Long = 0
    Public Property Temperature As Double = -999
    Public Property temp_trend As String = String.Empty
    Public Property Humidity As Integer = 0
    Public Property CO2 As Integer = 0
    Public Property date_max_temp As Long = -999
    Public Property date_min_temp As Long = -999
    Public Property min_temp As Double = -999
    Public Property max_temp As Double = -999
    ' Rain
    Public Property Rain As Double = 0
    Public Property sum_rain_24 As Double = 0
    Public Property sum_rain_1 As Double = 0
    ' Wind
    Public Property WindAngle As Integer = -999
    Public Property WindStrength As Integer = -999
    Public Property GustAngle As Integer = -999
    Public Property GustStrength As Integer = -999
    Public Property date_max_wind_str As Long = -999
    Public Property max_wind_angle As Integer = -999
    Public Property max_wind_str As Integer = -999
  End Class

  <Serializable()>
  Public Class DeviceDashboardData
    Public Property AbsolutePressure As Double = 0
    Public Property time_utc As Long = 0
    Public Property Noise As Integer = 0
    Public Property Temperature As Double = -999
    Public Property temp_trend As String = String.Empty
    Public Property Humidity As Integer = 0
    Public Property Pressure As Double = 0
    Public Property pressure_trend As String = String.Empty
    Public Property CO2 As Integer = 0
    Public Property date_max_temp As Long = -999
    Public Property date_min_temp As Long = -999
    Public Property min_temp As Double = -999
    Public Property max_temp As Double = -999
    Public Property WindAngle As Integer = 0
    Public Property WindStrength As Integer = 0
    Public Property GustAngle As Integer = 0
    Public Property GustStrength As Integer = 0
    Public Property date_max_wind_str As Long = 0
    Public Property max_wind_angle As Integer = 0
    Public Property max_wind_str As Integer = 0
  End Class

  <Serializable()>
  Public Class User
    Public Property mail As String = String.Empty
    Public Property administrative As UserOption
  End Class

  <Serializable()>
  Public Class UserOption
    Public Property country As String = String.Empty
    Public Property reg_locale As String = String.Empty ' user regional preferences (used for displaying date)
    Public Property lang As String = String.Empty
    Public Property unit As Integer = 0                ' 0 -> metric system, 1 -> imperial system
    Public Property windunit As Integer = 0            ' 0 -> kph, 1 -> mph, 2 -> ms, 3 -> beaufort, 4 -> knot
    Public Property pressureunit As Integer = 0        ' 0 -> mbar, 1 -> inHg, 2 -> mmHg
    Public Property feel_like_algo As Integer = 0      ' 0 -> humidex, 1 -> heat-index 
  End Class
#End Region

#Region "HomeCoachData"

  <Serializable()>
  Public Class HomeCoachData
    Public Property body As DevicesAndUser2
    Public Property status As String = String.Empty
    Public Property time_exec As Double
    Public Property time_server As Double
  End Class

  <Serializable()>
  Public Class DevicesAndUser2
    Public Property devices As List(Of Device2)
    Public Property user As User
  End Class

  <Serializable()>
  Public Class Device2
    Public Property _id As String = String.Empty
    Public Property cipher_id As String = String.Empty
    Public Property last_status_store As Integer = 0
    Public Property place As Place
    Public Property name As String = String.Empty
    Public Property type As String = String.Empty
    Public Property dashboard_data As DeviceDashboardData2    ' Last data measured per device
    Public Property data_type As String()                     ' Array of data measured by the device (e.g. "Temperature","Humidity")
    Public Property co2_calibrating As Boolean = False
    Public Property date_setup As Long = 0                    ' Date when the Weather station was set up
    Public Property last_setup As Long = 0                    ' Date when the Weather station was set up
    Public Property firmware As Long = 0                      ' Version of the software
    Public Property last_upgrade As Long = 0
    Public Property wifi_status As Integer = 0
  End Class

  <Serializable()>
  Public Class DeviceDashboardData2
    Public Property AbsolutePressure As Double = 0
    Public Property time_utc As Long = 0
    Public Property health_idx As Integer = 0
    Public Property Noise As Integer = 0
    Public Property Temperature As Double = -999
    Public Property temp_trend As String = String.Empty
    Public Property Humidity As Integer = 0
    Public Property Pressure As Double = 0
    Public Property pressure_trend As String = String.Empty
    Public Property CO2 As Integer = 0
    Public Property date_max_temp As Long = -999
    Public Property date_min_temp As Long = -999
    Public Property min_temp As Double = -999
    Public Property max_temp As Double = -999
  End Class

#End Region

End Class
