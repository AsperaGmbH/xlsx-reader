<?php
namespace Aspera\Spreadsheet\XLSX;

use XMLReader;
use InvalidArgumentException;

/**
 * Extension of XMLReader to ease parsing of XML files of the OOXML specification.
 *
 * Depending on edition, namespaceUris in OOXML documents can be entirely different. Besides adding extra
 * matching overhead, this makes custom-made documents that are employing their own namespace rules a bit
 * complicated to read correctly. To mitigate the impact of that, this wrapper of XMLReader supplies methods
 * that deal with these issues automatically.
 */
class OoxmlReader extends XMLReader
{
    /**
     * Identifiers of supported OOXML Namespaces.
     * Use these instead of namespaceUris when checking for elements that are part of OOXML namespaces.
     *
     * @var array NS_NONE Also known as the "empty" namespace. All attributes always default to this.
     * @var array XMLNS_XLSX_MAIN Root namespace of most XLSX documents.
     * @var array XMLNS_RELATIONSHIPS_DOCUMENTLEVEL Namespace used for references to relationship documents.
     * @var array XMLNS_RELATIONSHIPS_PACKAGELEVEL Root namespace used within relationship documents.
     */
    const NS_NONE = '';
    const NS_XLSX_MAIN = 'xlsx_main';
    const NS_RELATIONSHIPS_DOCUMENTLEVEL = 'relationships_documentlevel';
    const NS_RELATIONSHIPS_PACKAGELEVEL = 'relationships_packagelevel';

    /** @var array Format: $namespace_list[-XMLNS_IDENTIFIER-][-INTRODUCING_EDITION_OF_SPECIFICATION-] = -NAMESPACE_URI- */
    private $namespace_list;

    /** @var string One of the NS_ constants that will be used if methods requiring a NsId for an element tag do not get one delivered. */
    private $default_namespace_identifier_elements;

    /** @var string One of the NS_ constants that will be used if methods requiring a NsId for an attribute do not get one delivered. */
    private $default_namespace_identifier_attributes;

    public function __construct()
    {
        $this->initNamespaceList();
        // Note: No parent::__construct() - XMLReader does not have its own constructor.
    }

    /**
     * Initialize $this->namespace_list.
     */
    private function initNamespaceList()
    {
        $this->namespace_list = array(
            self::NS_NONE => array(''),
            self::NS_XLSX_MAIN => array(
                1 => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                3 => 'http://purl.oclc.org/ooxml/spreadsheetml/main'
            ),
            self::NS_RELATIONSHIPS_DOCUMENTLEVEL => array(
                1 => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                3 => 'http://purl.oclc.org/ooxml/officeDocument/relationships'
            ),
            self::NS_RELATIONSHIPS_PACKAGELEVEL  => array(
                1 => 'http://schemas.openxmlformats.org/package/2006/relationships',
                3 => 'http://purl.oclc.org/ooxml/officeDocument/relationships' // Note: Same as DOCUMENTLEVEL
            )
        );
    }

    /**
     * Sets the default namespace_identifier for element tags,
     * to be used when methods requiring a namespace_identifier are not given one.
     *
     * @param   string  $namespace_identifier
     *
     * @throws  InvalidArgumentException
     */
    public function setDefaultNamespaceIdentifierElements($namespace_identifier)
    {
        if (!isset($this->namespace_list[$namespace_identifier])) {
            throw new InvalidArgumentException('unknown namespace identifier [' . $namespace_identifier . ']');
        }
        $this->default_namespace_identifier_elements = $namespace_identifier;
    }

    /**
     * Sets the default namespace_identifier for element attributes,
     * to be used when methods requiring a namespace_identifier are not given one.
     *
     * @param   string  $namespace_identifier
     *
     * @throws  InvalidArgumentException
     */
    public function setDefaultNamespaceIdentifierAttributes($namespace_identifier)
    {
        if (!isset($this->namespace_list[$namespace_identifier])) {
            throw new InvalidArgumentException('unknown namespace identifier [' . $namespace_identifier . ']');
        }
        $this->default_namespace_identifier_attributes = $namespace_identifier;
    }

    /**
     * Checks if the element the reader is currently pointed at is of the given local_name with a namespace_uri
     * that matches the list of namespaces identified by the given namespace_identifier constant.
     *
     * @param   string      $local_name
     * @param   string|null $namespace_identifier   NULL = Fallback to $this->default_namespace_identifier_elements
     * @return  bool
     *
     * @throws  InvalidArgumentException
     */
    public function matchesElement($local_name, $namespace_identifier = null)
    {
        return $this->localName === $local_name
            && $this->matchesNamespace($namespace_identifier);
    }

    /**
     * Checks if any of the given list of elements is matched by the current element.
     * Returns the array key of the element that matched, or false if none matched.
     *
     * @param   array   $list_of_elements   Format: array([MATCH_1_ID] => array(LOCAL_NAME_1, NAMESPACE_ID_1), ...)
     * @return  mixed|false If no match was found: false. Otherwise, the parameter array's key of the element definition that matched.
     */
    public function matchesOneOfList($list_of_elements)
    {
        foreach ($list_of_elements as $one_element_key => $one_element) {
            $parameter_count = count($one_element);
            if ($parameter_count < 1 || $parameter_count > 2) {
                throw new InvalidArgumentException('Invalid definition of element. Expected 1 or 2 parameters, got [' . $parameter_count . '].');
            }
            if ($this->localName !== $one_element[0]) {
                continue;
            }
            if (!isset($one_element[1])) {
                $one_element[1] = null; // default $namespace_identifier value
            }
            if ($this->matchesNamespace($one_element[1])) {
                return $one_element_key;
            }
        }
        return false;
    }

    /**
     * Checks if the element the reader is currently pointed at contains an element with a namespace_uri
     * that matches the list of namespaces identified by the given namespace_identifier constant.
     *
     * @param   string|null $namespace_identifier   NULL = Fallback to $this->default_namespace_identifier_elements
     * @param   bool        $for_attribute          Determines the scope of validation; true: attribute, false: element tag
     * @return  bool
     *
     * @throws  InvalidArgumentException
     */
    public function matchesNamespace($namespace_identifier = null, $for_attribute = false)
    {
        return in_array(
            $this->namespaceURI,
            $this->namespace_list[$this->validateNamespaceIdentifier($namespace_identifier, $for_attribute)],
            true
        );
    }

    /**
     * Checks if the current element is a closing tag / END_ELEMENT.
     *
     * @return bool
     */
    public function isClosingTag() {
        return $this->nodeType === OoxmlReader::END_ELEMENT;
    }

    /**
     * Extension of getAttributeNs that checks with a namespace_identifier rather than a specific namespace_uri.
     *
     * @param   string      $local_name
     * @param   string|null $namespace_identifier   NULL = Fallback to $this->default_namespace_identifier_elements
     * @return  NULL|string
     *
     * @throws  InvalidArgumentException
     */
    public function getAttributeNsId($local_name, $namespace_identifier = null)
    {
        $namespace_identifier = $this->validateNamespaceIdentifier($namespace_identifier, true);

        $ret_value = null;
        foreach ($this->namespace_list[$namespace_identifier] as $namespace_uri) {
            $moved_successfully = ($namespace_uri === '')
                ? $this->moveToAttribute($local_name)
                : $this->moveToAttributeNs($local_name, $namespace_uri);
            if ($moved_successfully) {
                $ret_value = $this->value;
                break;
            }
        }
        $this->moveToElement();

        return $ret_value;
    }

    /**
     * Moves to the next node matching the given criteria.
     *
     * @param   string      $local_name
     * @param   string|null $namespace_identifier
     * @return  bool
     */
    public function nextNsId($local_name, $namespace_identifier = null)
    {
        while ($this->next($local_name)) {
            if ($this->matchesNamespace($namespace_identifier)) {
                return true;
            }
        }

        return false;
    }

    /**
     * Checks if the given namespace_identifier is valid. If null is given, will try to fallback to
     * $this->default_namespace_identifier_elements. Returns the correct namespace_identifier for further usage.
     *
     * @param   string|null $namespace_identifier
     * @param   bool        $for_attribute          Determines the default namespace_identifier to fallback to.
     * @return  string
     *
     * @throws  InvalidArgumentException
     */
    private function validateNamespaceIdentifier($namespace_identifier, $for_attribute = false)
    {
        if ($namespace_identifier === null) {
            $default_namespace_identifier = ($for_attribute)
                ? $this->default_namespace_identifier_attributes
                : $this->default_namespace_identifier_elements;
            if ($default_namespace_identifier === null) {
                throw new InvalidArgumentException('no namespace identifier given');
            }

            return $default_namespace_identifier;
        } elseif (!isset($this->namespace_list[$namespace_identifier])) {
            throw new InvalidArgumentException('unknown namespace identifier [' . $namespace_identifier . ']');
        }

        return $namespace_identifier;
    }
}
