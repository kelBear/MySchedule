<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl">
  <xsl:output method="html" omit-xml-declaration="yes" indent="yes"/>

  <xsl:template match="@* | node()">
        
    <html>
      <head>
        <style>
          .tableStyle {font-family: Arial; border: 1 solid #808080; font-size: 12;}
          .tableHeader {font-size: 14; background-color: #51B0F5; color:white}
        </style>
      </head>
      <body>
        <table class="tableStyle" width="80%">
          <tr class="tableHeader">
            <xsl:for-each select="/*/node()">
              <xsl:if test="position()=1">
                <xsl:for-each select="*">
                  <td style="padding:5; text-align:center">
                    <b><xsl:value-of select="local-name()"/></b>
                  </td>
                </xsl:for-each>
              </xsl:if>
            </xsl:for-each>
          </tr>
          <xsl:for-each select="*">
            <xsl:variable name="rowColour">
              <xsl:choose>
                <xsl:when test="position() mod 2 = 0">#D4EBFA</xsl:when>
                <xsl:otherwise>#BCE4F5</xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <tr bgcolor="{$rowColour}">
              <xsl:for-each select="*">
                <td style="padding-left:2; padding-right:5">
                    <xsl:value-of select="." />
                </td>
              </xsl:for-each>
            </tr>
          </xsl:for-each>
        </table>
        <br/>
      </body>
    </html>
  </xsl:template>
</xsl:stylesheet>