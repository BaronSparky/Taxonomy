<?xml version="1.0" encoding="utf-8"?>
<!-- 
    To use:
    
    1. Download the tool msxsl.exe from Microsoft to your hard drive (c:)
      [http://www.microsoft.com/downloads/details.aspx?FamilyId=2FB55371-C94E-4373-B0E9-DB4816552E41&displaylang=en]
    
    2. Open a command prompt using Start > Run and type cmd
    
      c:\msxsl.exe c:\<yyyymmdd>-RDM_Export.xml c:\SPO-SubjectCategory.xslt -o c:\<yyyymmdd>-SubjectCategory_RDM.xml
  -->

<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:hiq="http://vpsl.co.uk" exclude-result-prefixes="hiq">

  <xsl:output method="xml" indent="yes" />
  <xsl:strip-space elements="*" />
  <xsl:key name="object_key" match="hiq:CVRObject" use="@id"/>

  <xsl:template match="*">
    <xsl:apply-templates select="@*" />
    <xsl:apply-templates select="node()" />
  </xsl:template>

  <xsl:template match="@*">
  </xsl:template>

  <xsl:variable name="TopLevelNodes" select="/hiq:CVRObjects/hiq:MODXMLbodyStructure/hiq:CVRObject[@name and not(hiq:ObjectRelationships/hiq:ObjectRelationshipItem/hiq:ObjectRelationshipType/@name='Non Preferred Term For') and not(hiq:ObjectRelationships/hiq:ObjectRelationshipItem/hiq:ObjectRelationshipType/@name='Child Of')]" />
  <xsl:template match="/">
    <TermStores>
      <TermStore Name="" IsOnline="" WorkingLanguage="" DefaultLanguage="" SystemGroup="">
        <Groups>
          <Group Id="TRANSFORMED" Name="Subject Category (UK Defence Taxonomy)" Description="" IsSystemGroup="" IsSiteCollectionGroup="">
            <TermSets>
              <TermSet Id="" Name="Subject Category" Description="Use Subject Categories to describe, as specifically as possible, the document content. Selected from the UK Defence Taxonomy." IsAvailableForTagging="True" IsOpenForTermCreation="False" TermCount="">
                <CustomProperties>
                  <CustomProperty Key="" Value=""/>
                </CustomProperties>
                <Terms>
                  <xsl:for-each select="$TopLevelNodes">
                    <xsl:sort select="@name"/>
                    <!-- This is the Taxonomy view of the Defence structure -->
                    <xsl:if test="hiq:Attributes/hiq:TypeAttributes/hiq:IsTaxonomy[@state='true']">
                      <Term Name="{@name}" Id="" CustomSortOrder="" IsAvailableForTagging="" IsSourceTerm="" IsRoot="" IsReused="" IsKeyword="" IsDeprecated="" CVRObjectID="{@id}" Processed="False">
                        <Descriptions>
                          <!--We look for "Term Notes" and use that as the nodes "Description"-->
                          <xsl:variable name="Description" select="hiq:Attributes/hiq:TypeAttributes/hiq:TypeTextArea" />
                          <Description Language="1033" Value="{$Description}"/>
                        </Descriptions>
                        <CustomProperties>
                          <CustomProperty Key="" value="" />
                        </CustomProperties>
                        <LocalCustomProperties>
                          <LocalCustomProperty Key="CVRObjectID" Value="{@id}"/>
                        </LocalCustomProperties>
                        <Labels>
                          <Label Value="{@name}" Language="1033" IsDefaultForLanguage="True"/>
                          <!--The node may have a number of labels/non-preferred terms. Using a choose so we can do things with other terms later if needed-->
                          <xsl:for-each select="hiq:ObjectRelationships/hiq:ObjectRelationshipItem">
                            <xsl:sort select="hiq:ObjectRelationshipType/@name"/>
                            <xsl:choose>
                              <!-- Narrower Term relationships -->
                              <xsl:when test="hiq:ObjectRelationshipType/@name = 'Preferred Term For'">
                                <Label Value="{hiq:ObjectRelationshipObject/@name}" Language="1033" IsDefaultForLanguage="False"/>
                              </xsl:when>
                              <!-- Don't output other relationship types -->
                              <xsl:otherwise>
                              </xsl:otherwise>
                            </xsl:choose>
                          </xsl:for-each>
                        </Labels>
                        <Terms>
                          <!-- Loop around all the relationships -->
                          <xsl:for-each  select="hiq:ObjectRelationships/hiq:ObjectRelationshipItem[hiq:ObjectRelationshipType/@name='Parent Of']/hiq:ObjectRelationshipObject">
                          <xsl:sort select="@name"/>
                            <xsl:call-template name="node_children">
                              <xsl:with-param name="node_id" select="@id"/>
                              <xsl:with-param name="node_name" select="@name"/>
                              <xsl:with-param name="node_level" select="1"/>
                            </xsl:call-template>
                          </xsl:for-each>
                        </Terms>
                      </Term>
                    </xsl:if>
                  </xsl:for-each>
                </Terms>
              </TermSet>
            </TermSets>
          </Group>
        </Groups>
      </TermStore>
    </TermStores>
  </xsl:template>

  <xsl:template name="node_children">
    <!-- recursive loop until done -->
    <xsl:param name="node_id"/>
    <xsl:param name="node_name"/>
    <xsl:param name="node_level"/>
    <xsl:variable name="IsTaxonomy" select="key('object_key',$node_id)/hiq:Attributes/hiq:TypeAttributes/hiq:IsTaxonomy/@state" />
    <xsl:choose>
      <xsl:when test="$IsTaxonomy = 'true'">
        <Term Id="" Name="{$node_name}" IsDeprecated="" IsAvailableForTagging="" IsKeyword="" IsReused="" IsRoot="" IsSourceTerm="" CustomSortOrder="" CVRObjectID="{$node_id}" Processed="False">
          <Descriptions>
            <!--We look for "Term Notes" and use that as the nodes "Description"-->
            <xsl:variable name="Description" select="key('object_key',$node_id)/hiq:Attributes/hiq:TypeAttributes/hiq:TypeTextArea" />
            <Description Language="1033" Value="{$Description}" />
          </Descriptions>
          <CustomProperties>
            <CustomProperty Key="" Value="" />
          </CustomProperties>
          <LocalCustomProperties>
            <LocalCustomProperty Key="CVRObjectID" Value="{$node_id}" />
          </LocalCustomProperties>
          <Labels>
            <Label Value="{$node_name}" Language="1033" IsDefaultForLanguage="True"/>
            <!--The node may have a number of labels/non-preferred terms-->
            <xsl:for-each select="key('object_key',$node_id)/hiq:ObjectRelationships/hiq:ObjectRelationshipItem[hiq:ObjectRelationshipType/@id=1]/hiq:ObjectRelationshipObject">
              <xsl:sort select="@name"/>
              <xsl:variable name="Label" select="@name" />
              <Label Value="{$Label}" Language="1033" IsDefaultForLanguage="False"/>
            </xsl:for-each>
          </Labels>
          <Terms>
            <xsl:if test="$node_level &lt; 6">
              <xsl:for-each select="key('object_key',$node_id)/hiq:ObjectRelationships/hiq:ObjectRelationshipItem[hiq:ObjectRelationshipType/@name='Parent Of']/hiq:ObjectRelationshipObject">
                <xsl:sort select="@name"/>
                <xsl:call-template name="node_children">
                  <xsl:with-param name="node_id" select="@id"/>
                  <xsl:with-param name="node_name" select="@name"/>
                  <xsl:with-param name="node_level" select="$node_level + 1"/>
                </xsl:call-template>
              </xsl:for-each>
            </xsl:if>
          </Terms>
        </Term>
      </xsl:when>
      <xsl:otherwise>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="comment() | processing-instruction() | text()">
    <xsl:copy />
  </xsl:template>

</xsl:stylesheet>
