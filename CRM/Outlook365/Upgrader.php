<?php

declare(strict_types = 1);

/**
 * Collection of upgrade steps.
 */
class CRM_Outlook365_Upgrader extends CRM_Extension_Upgrader_Base {

  /**
   * An uninstalled earlier version leaves its data processors behind; install skips upgrade steps.
   */
  public function install(): void {
    self::removeDataProcessors();
  }

  public function upgrade_3200(): bool {
    self::removeDataProcessors();
    return TRUE;
  }

  /**
   * Removes the contact and group data processors earlier versions installed.
   * Their tables belong to the dataprocessor extension; child rows go by ON DELETE CASCADE.
   */
  private static function removeDataProcessors(): void {
    if (!CRM_Core_DAO::checkTableExists('civicrm_data_processor')) {
      return;
    }
    CRM_Core_DAO::executeQuery(
      "DELETE FROM civicrm_data_processor WHERE name IN ('outlook_365_contacts', 'outlook_365_groups')"
    );
    if (class_exists(CRM_Dataprocessor_Utils_Cache::class)) {
      CRM_Dataprocessor_Utils_Cache::clearAllDataProcessorCaches();
    }
  }

}
